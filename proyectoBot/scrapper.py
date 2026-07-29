import os
import sys
import time
import re
import json
import gspread
import requests
import smtplib
import concurrent.futures
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from oauth2client.service_account import ServiceAccountCredentials
from urllib.parse import urlparse
from bs4 import BeautifulSoup
from typing import List, Dict, Set, Optional, Tuple

# Try to load local .env file if python-dotenv is available
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# Reconfigure stdout and stderr to UTF-8 to prevent UnicodeEncodeError in Windows terminals
try:
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except (AttributeError, ValueError):
    pass

# --- CONSTANTS & CONFIGURATION ---
SCOPE = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1hW8G3BxoUN3f9ufaKjHHbu_Uh0HD4rVBclRfVVrpr5E/edit?usp=sharing"
DEFAULT_CREDS_FILE = "credenciales.json"

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}

# Suffixes to append to each company portal url for seeking jobs
SEARCH_SUFFIXES = ["/trabajos", "/analista", ""]

# Job alert keywords (case-insensitive)
ALERT_KEYWORDS = ["analista", "consultor"]

# Max worker threads for scraping in parallel
MAX_WORKERS = 5


# --- GOOGLE SHEETS UTILITIES ---

def get_gspread_client() -> gspread.Client:
    """
    Initializes the Google Sheets client using a credentials file.
    Checks environment variable, local directory, and script directory.
    """
    # 1. Check environment variable for credentials filepath
    creds_path = os.getenv("GOOGLE_CREDS_PATH", DEFAULT_CREDS_FILE)
    
    # 2. If it doesn't exist, check script's directory as fallback
    if not os.path.exists(creds_path):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        creds_path = os.path.join(script_dir, DEFAULT_CREDS_FILE)

    if not os.path.exists(creds_path):
        raise FileNotFoundError(
            f"Google credentials file not found at: {creds_path}. "
            "Please create it or configure GOOGLE_CREDS_PATH."
        )

    creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, SCOPE)
    return gspread.authorize(creds)


def read_portales(client: gspread.Client, sheet_url: str) -> List[Tuple[str, str]]:
    """
    Reads company names and URLs from the 'Portales' tab (Column A and B).
    """
    try:
        sheet = client.open_by_url(sheet_url).worksheet("Portales")
        all_values = sheet.get_all_values()
        if not all_values:
            return []
        
        # Skip header row (index 0) and filter out empty rows
        portales = []
        for row in all_values[1:]:
            if len(row) >= 2 and row[0].strip() and row[1].strip():
                portales.append((row[0].strip(), row[1].strip()))
        return portales
    except Exception as e:
        print(f"[-] Error reading 'Portales' sheet: {e}")
        return []


def read_historial(client: gspread.Client, sheet_url: str) -> Set[str]:
    """
    Reads already-sent offer URLs from the 'Historial' tab (Column A) to prevent duplicates.
    """
    try:
        sheet = client.open_by_url(sheet_url).worksheet("Historial")
        all_values = sheet.get_all_values()
        if not all_values:
            return set()
        
        # We assume Column A contains the URL. Filter out headers or empty values.
        historial = set()
        for row in all_values:
            if row and row[0].strip():
                historial.add(row[0].strip().lower())
        return historial
    except Exception as e:
        print(f"[-] Error reading 'Historial' sheet: {e}")
        # Return empty set to avoid breaking the script
        return set()


def append_to_historial(client: gspread.Client, sheet_url: str, urls: List[str]) -> None:
    """
    Appends newly sent job URLs to the 'Historial' sheet using a single batch update.
    """
    if not urls:
        return
    try:
        sheet = client.open_by_url(sheet_url).worksheet("Historial")
        # Prepare rows to append: each URL goes in Column A
        rows = [[url] for url in urls]
        sheet.append_rows(rows)
        print(f"[+] Successfully registered {len(urls)} URLs in the 'Historial' sheet.")
    except Exception as e:
        print(f"[-] Error updating 'Historial' sheet: {e}")


# --- NU3X DEVALUE DECODER (FOR TRABAJANDO.CL PORTALS) ---

def resolve_nuxt_data(data: list) -> List[dict]:
    """
    Decodes the Nuxt 3 devalue payload format. Nuxt 3 serializes object states
    where values inside dicts/lists point to indices in the same flat array.
    """
    resolved_cache = {}

    def get_val(idx):
        if not isinstance(idx, int) or idx < 0 or idx >= len(data):
            return idx
        if idx in resolved_cache:
            return resolved_cache[idx]
        
        # Guard against infinite recursion
        resolved_cache[idx] = f"<recursive_ref_{idx}>"
        
        raw_val = data[idx]
        if isinstance(raw_val, dict):
            resolved = {}
            for k, v in raw_val.items():
                resolved[k] = get_val(v)
            resolved_cache[idx] = resolved
            return resolved
        elif isinstance(raw_val, list):
            resolved = [get_val(x) for x in raw_val]
            resolved_cache[idx] = resolved
            return resolved
        else:
            resolved_cache[idx] = raw_val
            return raw_val

    jobs = []
    for idx, item in enumerate(data):
        if isinstance(item, dict) and 'idOferta' in item and 'nombreCargo' in item:
            resolved_job = get_val(idx)
            # Filter templates/empty offers in Nuxt payload
            if resolved_job.get('idOferta') and resolved_job.get('nombreCargo'):
                jobs.append(resolved_job)
    return jobs


# --- SCRAPING LOGIC ---

def fetch_url_with_retry(url: str, max_retries: int = 3) -> Optional[str]:
    """
    Fetches the HTML content of a URL with exponential backoff retries on HTTP 429 / connection issues.
    """
    delay = 2.0
    for attempt in range(max_retries):
        try:
            response = requests.get(url, headers=HEADERS, timeout=12)
            if response.status_code == 200:
                return response.text
            elif response.status_code == 429:
                print(f"[!] HTTP 429 (Too Many Requests) for {url}. Retrying in {delay}s...")
                time.sleep(delay)
                delay *= 2
            elif response.status_code == 403:
                # 403 usually means cloudflare / security block, retry once after sleep
                print(f"[!] HTTP 403 (Forbidden) for {url}. Retrying in {delay}s...")
                time.sleep(delay)
                delay *= 2
            else:
                # E.g. 404, no need to retry
                return None
        except requests.exceptions.RequestException as e:
            print(f"[!] Connection error for {url} (Attempt {attempt+1}/{max_retries}): {e}")
            time.sleep(delay)
            delay *= 2
    return None


def scrape_portal(company_name: str, base_url: str) -> List[Dict]:
    """
    Scrapes a portal using multiple search suffixes and fallback strategies.
    Returns list of dicts with keys: title, url, company, date_text.
    """
    scraped_offers = []
    seen_urls = set()
    
    # Clean the base URL
    base_url = base_url.strip().rstrip('/')
    
    # Try all suffixes (plus base URL fallback)
    for suffix in SEARCH_SUFFIXES:
        target_url = f"{base_url}{suffix}"
        html = fetch_url_with_retry(target_url)
        if not html:
            continue
            
        soup = BeautifulSoup(html, 'html.parser')
        
        # --- Strategy 1: Nuxt Data Payload ---
        nuxt_script = soup.find('script', id='__NUXT_DATA__')
        nuxt_jobs_found = False
        if nuxt_script and nuxt_script.string:
            try:
                raw_data = json.loads(nuxt_script.string)
                jobs = resolve_nuxt_data(raw_data)
                if jobs:
                    nuxt_jobs_found = True
                    for job in jobs:
                        offer_id = job.get('idOferta')
                        cargo = job.get('nombreCargo', '').strip()
                        pub_date = job.get('publicadoHace', '') or job.get('fechaPublicacion', '')
                        empresa = job.get('nombreEmpresa', company_name).strip()
                        offer_url = f"{base_url}/trabajo/{offer_id}"
                        
                        if offer_url.lower() not in seen_urls:
                            seen_urls.add(offer_url.lower())
                            scraped_offers.append({
                                'title': cargo,
                                'url': offer_url,
                                'company': empresa,
                                'date_text': str(pub_date).strip()
                            })
            except Exception as e:
                print(f"[-] Nuxt parsing error for {company_name} at {target_url}: {e}")
                
        # --- Strategy 2: Fallback HTML Scraping ---
        # If Nuxt script was missing or failed to parse jobs, scan normal HTML links
        if not nuxt_jobs_found:
            for a in soup.find_all('a', href=True):
                href = a['href'].strip()
                if href.startswith('/'):
                    href = f"{base_url}{href}"
                    
                # Normalize relative URLs and search for job-like paths
                is_job_link = any(x in href.lower() for x in ['/ofertas/', '/empleos/', '/empleo/', '/oferta/', '/trabajo/', '/detalle/'])
                if is_job_link:
                    title = a.get_text(strip=True)
                    if not title:
                        continue
                    
                    # Try to capture date or context from parent elements
                    parent_text = a.parent.get_text(strip=True) if a.parent else ''
                    date_text = "hoy" if "hoy" in parent_text.lower() else ""
                    
                    if href.lower() not in seen_urls:
                        seen_urls.add(href.lower())
                        scraped_offers.append({
                            'title': title,
                            'url': href,
                            'company': company_name,
                            'date_text': date_text
                        })
                        
    return scraped_offers


# --- EMAIL NOTIFICATION SYSTEM ---

def send_html_email(recipient_email: str, gmail_user: str, gmail_pass: str, offers: List[Dict]) -> None:
    """
    Dispatches a high-quality HTML email listing new job offers using smtplib.
    """
    if not offers:
        return

    # Build the HTML content for the jobs list
    jobs_list_html = ""
    for idx, offer in enumerate(offers, 1):
        # Humanize dates slightly
        date_display = offer['date_text'] if offer['date_text'] else "No especificada"
        
        jobs_list_html += f"""
        <div class="job-card">
            <div class="job-title">{idx}. {offer['title']}</div>
            <div class="job-meta">
                <span><strong>Empresa:</strong> {offer['company']}</span>
                <span><strong>Publicado:</strong> {date_display}</span>
            </div>
            <a href="{offer['url']}" target="_blank" class="btn-apply">Ver Oferta Directa</a>
        </div>
        """

    current_time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Full HTML email document
    html_body = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Nuevas Ofertas de Trabajo</title>
        <style>
            body {{
                font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
                background-color: #f3f4f6;
                margin: 0;
                padding: 20px;
                color: #1f2937;
            }}
            .container {{
                max-width: 600px;
                margin: 0 auto;
                background-color: #ffffff;
                border-radius: 12px;
                overflow: hidden;
                box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
                border: 1px solid #e5e7eb;
            }}
            .header {{
                background: linear-gradient(135deg, #4f46e5 0%, #7c3aed 100%);
                padding: 30px 20px;
                text-align: center;
                color: #ffffff;
            }}
            .header h1 {{
                margin: 0;
                font-size: 24px;
                font-weight: 700;
                letter-spacing: -0.025em;
            }}
            .header p {{
                margin: 8px 0 0 0;
                font-size: 14px;
                color: #e0e7ff;
            }}
            .content {{
                padding: 24px;
                background-color: #ffffff;
            }}
            .job-card {{
                border: 1px solid #e5e7eb;
                border-radius: 8px;
                padding: 18px;
                margin-bottom: 16px;
                background-color: #ffffff;
            }}
            .job-title {{
                font-size: 18px;
                font-weight: 600;
                color: #111827;
                margin: 0 0 6px 0;
            }}
            .job-meta {{
                font-size: 13px;
                color: #6b7280;
                margin-bottom: 14px;
            }}
            .job-meta span {{
                margin-right: 16px;
            }}
            .btn-apply {{
                display: inline-block;
                background-color: #4f46e5;
                color: #ffffff !important;
                text-decoration: none;
                padding: 8px 16px;
                border-radius: 6px;
                font-size: 13px;
                font-weight: 600;
                text-align: center;
            }}
            .footer {{
                background-color: #f9fafb;
                padding: 16px 20px;
                text-align: center;
                font-size: 12px;
                color: #9ca3af;
                border-top: 1px solid #e5e7eb;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>Nuevas Ofertas Encontradas</h1>
                <p>Se encontraron {len(offers)} ofertas laborales que coinciden con tus filtros de búsqueda.</p>
            </div>
            <div class="content">
                {jobs_list_html}
            </div>
            <div class="footer">
                <p>Este es un reporte automático enviado por el Bot de Scraping.</p>
                <p>Hora de ejecución: {current_time_str} CLT</p>
            </div>
        </div>
    </body>
    </html>
    """

    # Setup the MIME message
    msg = MIMEMultipart('alternative')
    msg['Subject'] = f"Alerta de Trabajo: {len(offers)} nuevas ofertas de Analista / Consultor"
    msg['From'] = gmail_user
    msg['To'] = recipient_email
    
    # Attach HTML body
    msg.attach(MIMEText(html_body, 'html'))

    print("[*] Connecting to SMTP server (smtp.gmail.com:587) to send alert...")
    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.ehlo()
        server.starttls()  # Upgrade connection to secure TLS
        server.ehlo()
        server.login(gmail_user, gmail_pass)
        server.sendmail(gmail_user, recipient_email, msg.as_string())
        server.quit()
        print(f"[+] Email notification successfully sent to {recipient_email}.")
    except Exception as e:
        print(f"[-] Failed to send email alert: {e}")


# --- PORTALES RE-POPULATION SITEMAP WORKER (ORIGINAL FUNCTIONER) ---

def fetch_subdomain_original(url: str) -> Optional[dict]:
    """
    Original helper to extract subdomains.
    """
    try:
        r = requests.get(url, headers=HEADERS, timeout=10)
        if r.status_code == 200:
            soup = BeautifulSoup(r.text, 'html.parser')
            for a in soup.find_all('a', href=True):
                href = a['href']
                if 'trabajando.cl' in href:
                    match = re.search(r'https?://([a-zA-Z0-9.-]+)\.trabajando\.cl', href)
                    if match:
                        sub = match.group(1).lower()
                        if sub not in ['www', 'ayuda', 'personas', 'empresas', 'staticcdn']:
                            h1 = soup.find('h1')
                            company_name = h1.text.strip() if h1 else sub.capitalize()
                            return {
                                'subdomain': sub,
                                'url': f"https://{sub}.trabajando.cl",
                                'name': company_name
                            }
    except Exception:
        pass
    return None


def populate_portales_from_sitemap(client: gspread.Client, sheet_url: str, max_resultados: int = 1000) -> None:
    """
    Original logic to fetch sitemap and update the Portales sheet.
    """
    print("[*] Fetching company list from sitemap-empresas.xml...")
    sitemap_url = "https://www.trabajando.cl/sitemap-empresas.xml"
    
    try:
        r = requests.get(sitemap_url, headers=HEADERS, timeout=15)
        if r.status_code in [403, 429]:
            print("\n[!] ERROR: Access denied (403/429) by sitemap host.")
            return
        elif r.status_code != 200:
            print(f"[-] Error reading sitemap. Code: {r.status_code}")
            return
    except Exception as e:
        print(f"[-] Connection error with sitemap: {e}")
        return
        
    soup = BeautifulSoup(r.text, 'xml')
    urls = [u.text for u in soup.find_all('loc')]
    total_urls = len(urls)
    print(f"[+] Found {total_urls} empresas. Analyzing in parallel...")
    
    subdominios_vistos = set()
    rows = []
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        future_to_url = {executor.submit(fetch_subdomain_original, url): url for url in urls}
        
        for idx, future in enumerate(concurrent.futures.as_completed(future_to_url)):
            res = future.result()
            if res:
                sub = res['subdomain']
                if sub not in subdominios_vistos:
                    subdominios_vistos.add(sub)
                    rows.append([res['name'], res['url']])
                    print(f"[{len(rows)}] Found: {res['name']} -> {res['url']}")
                    
                    if len(rows) >= max_resultados:
                        for f in future_to_url:
                            f.cancel()
                        break
            
            if (idx + 1) % 50 == 0:
                print(f"Progress: {idx + 1}/{total_urls} pages analyzed...")
                
    if not rows:
        print("[-] No portals found or IP blocked.")
        return
        
    print(f"\n[+] Found {len(rows)} portals. Writing to sheet...")
    try:
        sheet = client.open_by_url(sheet_url).worksheet("Portales")
        sheet.clear()
        sheet.update(range_name='A1', values=[['Empresa', 'Links']] + rows)
        print("[+] 'Portales' tab updated successfully!")
    except Exception as e:
        print(f"[-] Error writing to sheets: {e}")


# --- MAIN PIPELINE ---

def run_scraper_pipeline() -> None:
    """
    Main job alert scraping automation execution pipeline.
    """
    print("[*] Starting scraping and automation pipeline...")
    
    sheet_url = os.getenv("GOOGLE_SHEET_URL", DEFAULT_SHEET_URL)
    gmail_user = os.getenv("GMAIL_USER")
    gmail_pass = os.getenv("GMAIL_PASS")
    
    # Simple check on GMAIL credentials, output warning if not defined
    if not gmail_user or not gmail_pass:
        print("[WARNING] GMAIL_USER or GMAIL_PASS environment variables are not set.")
        print("Email alerts will be skipped. Set these variables to enable emails.")
        
    try:
        client = get_gspread_client()
    except Exception as e:
        print(f"[-] Failed to initialize Google Sheets client: {e}")
        sys.exit(1)

    print("[*] Retrieving company portals from Google Sheets...")
    portales = read_portales(client, sheet_url)
    if not portales:
        print("[-] No company portals found in 'Portales' sheet. Exiting.")
        sys.exit(0)
    print(f"[+] Found {len(portales)} portals to scrape.")

    print("[*] Retrieving history of already sent job offers...")
    historial_urls = read_historial(client, sheet_url)
    print(f"[+] Retrieved {len(historial_urls)} URLs from 'Historial' sheet.")

    print("[*] Initiating scraping tasks in parallel (max 5 workers)...")
    all_scraped_offers = []
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        future_to_portal = {
            executor.submit(scrape_portal, company, url): (company, url) 
            for company, url in portales
        }
        for future in concurrent.futures.as_completed(future_to_portal):
            company, url = future_to_portal[future]
            try:
                res = future.result()
                all_scraped_offers.extend(res)
                print(f"[+] Scraped {company} ({url}): {len(res)} raw offers found.")
            except Exception as e:
                print(f"[-] Error scraping {company} at {url}: {e}")

    # Deduplicate raw offers by URL
    unique_offers = {}
    for offer in all_scraped_offers:
        url_normalized = offer['url'].lower().strip()
        if url_normalized not in unique_offers:
            unique_offers[url_normalized] = offer

    print(f"\n[*] Extracted {len(unique_offers)} unique offers in total. Filtering matches...")
    
    nuevas_ofertas = []
    new_urls_to_log = []
    
    for url, offer in unique_offers.items():
        # Check against Historial
        if url in historial_urls:
            continue
            
        title_lower = offer['title'].lower()
        date_text_lower = offer['date_text'].lower()
        
        # Criteria matches if:
        # Title contains "analista" OR "consultor" OR date text contains "hoy"
        keyword_match = any(kw in title_lower for kw in ALERT_KEYWORDS)
        date_match = "hoy" in date_text_lower
        
        if keyword_match or date_match:
            nuevas_ofertas.append(offer)
            new_urls_to_log.append(offer['url'])  # Keep exact original casing for history log

    print(f"[+] Found {len(nuevas_ofertas)} matching NEW offers.")

    if nuevas_ofertas:
        # 1. Send email alerts if configuration is present
        if gmail_user and gmail_pass:
            send_html_email(
                recipient_email=gmail_user, 
                gmail_user=gmail_user, 
                gmail_pass=gmail_pass, 
                offers=nuevas_ofertas
            )
        else:
            print("[*] Skipping email dispatch because credentials are not configured.")
            print("New jobs found:")
            for idx, job in enumerate(nuevas_ofertas, 1):
                print(f"  {idx}. {job['title']} - {job['company']} -> {job['url']}")
        
        # 2. Append new URLs to Historial sheet
        print("[*] Logging new URLs to 'Historial' sheet...")
        append_to_historial(client, sheet_url, new_urls_to_log)
    else:
        print("[+] No new matches found. No emails sent and no history updates needed.")
        
    print("[+] Scraper pipeline completed successfully.")


if __name__ == "__main__":
    # Check for CLI arguments
    command = "run-scraper"
    if len(sys.argv) > 1:
        command = sys.argv[1]

    if command == "populate-portales":
        try:
            g_client = get_gspread_client()
            sheet_url = os.getenv("GOOGLE_SHEET_URL", DEFAULT_SHEET_URL)
            populate_portales_from_sitemap(g_client, sheet_url)
        except Exception as e:
            print(f"[-] Error: {e}")
    else:
        # Default behavior: run-scraper
        run_scraper_pipeline()
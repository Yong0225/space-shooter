#!/usr/bin/env python3
"""
Las Vegas Chinatown (Spring Mountain Rd) Nevada Leads Scraper
Collects restaurant / food business leads in LV Chinatown & surrounding area, NV
Output: Chinatown leads.xlsx  (Name | Area | Website | Email | Instagram | Facebook)

Resume-safe: writes to Excel + chinatown_lv_progress.json after EVERY lead.
Run: py scrap_chinatown_lv.py
      py scrap_chinatown_lv.py --reset   # clear progress and restart
"""

import re, json, time, random, sys, argparse, io
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# ── Config ────────────────────────────────────────────────────────────────────
TARGET      = 9999
OUTPUT      = "Chinatown leads.xlsx"
PROGRESS    = "chinatown_lv_progress.json"
HEADLESS    = False

SEARCH_QUERIES = [
    # Generic — Chinatown / Spring Mountain Rd
    "restaurant Chinatown Las Vegas Nevada",
    "food Chinatown Las Vegas Nevada",
    "cafe Chinatown Las Vegas Nevada",
    "restaurant Spring Mountain Road Las Vegas Nevada",
    "food Spring Mountain Road Las Vegas Nevada",
    "restaurant Spring Mountain Rd Las Vegas NV",
    "food Spring Mountain Rd Las Vegas NV",
    # Spring Mountain Rd by cross street
    "restaurant Spring Mountain Decatur Las Vegas Nevada",
    "food Spring Mountain Valley View Las Vegas Nevada",
    "restaurant Spring Mountain Arville Las Vegas Nevada",
    "food Spring Mountain Jones Las Vegas Nevada",
    "restaurant Spring Mountain Rainbow Las Vegas Nevada",
    "food Spring Mountain Blvd Las Vegas Nevada",
    # Valley View Blvd
    "restaurant Valley View Blvd Las Vegas Nevada",
    "food Valley View Las Vegas Nevada",
    # Arville St / Jones Blvd
    "restaurant Arville Street Las Vegas Nevada",
    "food Arville Las Vegas Nevada",
    "restaurant Jones Blvd Las Vegas Nevada",
    "food Jones Boulevard Las Vegas Nevada",
    # Decatur Blvd (western Chinatown)
    "restaurant Decatur Blvd Las Vegas Nevada",
    "food Decatur Boulevard Las Vegas Nevada",
    # Twain Ave (north of Spring Mountain)
    "restaurant Twain Avenue Las Vegas Nevada",
    "food Twain Ave Las Vegas Nevada",
    # Flamingo Rd near Chinatown
    "restaurant Flamingo Road Chinatown Las Vegas Nevada",
    "food W Flamingo Las Vegas Nevada",
    # Desert Inn Rd (south of Spring Mountain)
    "restaurant Desert Inn Road Las Vegas Nevada",
    "food W Desert Inn Las Vegas Nevada",
    # Asian cuisines — this is the heart of LV Asian food
    "Chinese restaurant Chinatown Las Vegas Nevada",
    "Chinese food Spring Mountain Las Vegas Nevada",
    "Chinese BBQ Las Vegas Nevada",
    "Dim sum Las Vegas Nevada",
    "Dim sum Spring Mountain Las Vegas Nevada",
    "Hong Kong style restaurant Las Vegas Nevada",
    "Cantonese restaurant Las Vegas Nevada",
    "Sichuan restaurant Las Vegas Nevada",
    "Shanghainese restaurant Las Vegas Nevada",
    "Taiwanese restaurant Las Vegas Nevada",
    "Taiwanese food Spring Mountain Las Vegas Nevada",
    "Boba Las Vegas Chinatown Nevada",
    "Bubble tea Spring Mountain Las Vegas Nevada",
    "Hot pot Las Vegas Nevada Chinatown",
    "KBBQ Las Vegas Nevada",
    "Korean BBQ Las Vegas Chinatown Nevada",
    "Korean restaurant Las Vegas Nevada Chinatown",
    "Korean food Spring Mountain Las Vegas Nevada",
    "Vietnamese restaurant Chinatown Las Vegas Nevada",
    "Vietnamese food Spring Mountain Las Vegas Nevada",
    "Pho Las Vegas Chinatown Nevada",
    "Banh mi Las Vegas Nevada",
    "Japanese restaurant Chinatown Las Vegas Nevada",
    "Sushi Spring Mountain Las Vegas Nevada",
    "Ramen Las Vegas Chinatown Nevada",
    "Izakaya Las Vegas Nevada",
    "Omakase Las Vegas Nevada",
    "Thai restaurant Chinatown Las Vegas Nevada",
    "Thai food Spring Mountain Las Vegas Nevada",
    "Filipino restaurant Las Vegas Nevada Chinatown",
    "Filipino food Spring Mountain Las Vegas Nevada",
    "Indian restaurant Chinatown Las Vegas Nevada",
    "Indian food Spring Mountain Las Vegas Nevada",
    "Malaysian restaurant Las Vegas Nevada",
    "Singapore food Las Vegas Nevada",
    "Indonesian restaurant Las Vegas Nevada",
    "Burmese restaurant Las Vegas Nevada",
    "Cambodian restaurant Las Vegas Nevada",
    "Laotian restaurant Las Vegas Nevada",
    "Mongolian BBQ Las Vegas Nevada",
    "Asian fusion Las Vegas Chinatown Nevada",
    "Dessert Chinatown Las Vegas Nevada",
    "Asian bakery Las Vegas Nevada",
    "Shaved ice Las Vegas Nevada",
    "Mochi Las Vegas Nevada",
    "Japanese bakery Las Vegas Nevada",
    "Matcha Las Vegas Nevada",
    # Non-Asian cuisines in the area
    "Mexican restaurant Spring Mountain Las Vegas Nevada",
    "Taqueria Spring Mountain Las Vegas Nevada",
    "Mediterranean Spring Mountain Las Vegas Nevada",
    "Pizza Spring Mountain Las Vegas Nevada",
    "Burger Spring Mountain Las Vegas Nevada",
    "Breakfast Spring Mountain Las Vegas Nevada",
    "Coffee shop Spring Mountain Las Vegas Nevada",
    # ZIP codes covering Chinatown / Spring Mountain corridor
    "restaurant Las Vegas Nevada 89102",
    "restaurant Las Vegas Nevada 89103",
    "restaurant Las Vegas Nevada 89147",
    "food Las Vegas Nevada 89102",
    "food Las Vegas Nevada 89103",
    "cafe Las Vegas Nevada 89102",
    "food Las Vegas Nevada 89147",
    "restaurant Las Vegas Nevada 89117 Spring Mountain",
]

BAD_EMAIL_DOMAINS = {
    'sentry.io', 'example.com', 'wixpress.com', 'squarespace.com',
    'wordpress.com', 'godaddy.com', 'shopify.com', 'amazonaws.com',
    'googleapis.com', 'schema.org', 'wix.com', 'weebly.com',
    'cloudflare.com', 'mailchimp.com', 'constantcontact.com',
}
BAD_EMAIL_PREFIXES = {'noreply', 'no-reply', 'donotreply', 'info@example', 'test@'}

ALCOHOL_SKIP_KEYWORDS = [
    'cocktail bar', 'cocktail lounge', 'cocktail room',
    'wine bar', 'wine lounge',
    'whiskey bar', 'whiskey lounge', 'bourbon bar',
    'spirits bar', 'spirits lounge',
    'speakeasy',
]

# ── Helpers ───────────────────────────────────────────────────────────────────
def is_alcohol_only(name):
    n = name.lower()
    return any(kw in n for kw in ALCOHOL_SKIP_KEYWORDS)

def sleep(mn=1.2, mx=2.8):
    time.sleep(random.uniform(mn, mx))

def ts():
    return time.strftime("%H:%M:%S")

# ── Progress ──────────────────────────────────────────────────────────────────
def load_progress():
    p = Path(PROGRESS)
    if p.exists():
        try:
            with open(p, encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {"query_idx": 0, "places": [], "done_urls": [], "leads": []}

def save_progress(prog):
    tmp = PROGRESS + ".tmp"
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(prog, f, ensure_ascii=False, indent=2)
    Path(tmp).replace(PROGRESS)

# ── Excel ─────────────────────────────────────────────────────────────────────
HEADERS  = ["Name", "Area", "Website", "Email", "Instagram", "Facebook"]
HDR_BG   = "2E4057"
ROW_BG   = ["FFFFFF", "EDF2F7"]
EMAIL_BG = "C6F6D5"

SEEN_FILE       = "scraped_seen.json"
_BOOTSTRAP_DIRS = ["已发送的leads", "未处理的leads", "leads_output", "."]

def _read_excels_for_seen(dirs):
    names, emails, scanned = set(), set(), []
    for d in dirs:
        dp = Path(d)
        if not dp.exists():
            continue
        for xlsx in dp.glob("*.xlsx"):
            try:
                wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
                ws = wb.active
                headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
                name_col = email_col = None
                for i, h in enumerate(headers):
                    if h and str(h).strip().lower() == "name":
                        name_col = i
                    if h and str(h).strip().lower() == "email":
                        email_col = i
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if not row:
                        continue
                    if name_col is not None and row[name_col]:
                        names.add(str(row[name_col]).lower().strip())
                    if email_col is not None and row[email_col]:
                        e = str(row[email_col]).lower().strip()
                        if "@" in e:
                            emails.add(e)
                wb.close()
                scanned.append(str(xlsx))
            except Exception:
                continue
    return names, emails, scanned

def _save_seen_db(names, emails):
    tmp = SEEN_FILE + ".tmp"
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump({"names": sorted(names), "emails": sorted(emails)}, f,
                  ensure_ascii=False, indent=2)
    Path(tmp).replace(SEEN_FILE)

def load_seen_db():
    p = Path(SEEN_FILE)
    if p.exists():
        try:
            with open(p, encoding='utf-8') as f:
                data = json.load(f)
            names  = set(data.get("names",  []))
            emails = set(data.get("emails", []))
            print(f"[SeenDB] {len(names)} names | {len(emails)} emails — loaded from {SEEN_FILE}")
            return names, emails
        except Exception:
            pass

    print(f"[SeenDB] {SEEN_FILE} not found — bootstrapping from existing lead files (one-time)...")
    names, emails, scanned = _read_excels_for_seen(_BOOTSTRAP_DIRS)
    print(f"[SeenDB] Bootstrapped from {len(scanned)} file(s): {len(names)} names | {len(emails)} emails")
    _save_seen_db(names, emails)
    print(f"[SeenDB] Saved to {SEEN_FILE} — future runs will use this file only")
    return names, emails

def load_existing_names():
    p = Path(OUTPUT)
    if not p.exists():
        return set()
    try:
        wb = openpyxl.load_workbook(OUTPUT, read_only=True)
        ws = wb.active
        names = set()
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and row[0]:
                names.add(str(row[0]).lower().strip())
        wb.close()
        return names
    except Exception:
        return set()

def count_existing_rows():
    p = Path(OUTPUT)
    if not p.exists():
        return 0
    try:
        wb = openpyxl.load_workbook(OUTPUT, read_only=True)
        ws = wb.active
        count = ws.max_row - 1
        wb.close()
        return max(count, 0)
    except Exception:
        return 0

def init_excel():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Chinatown Leads"
    ws.append(HEADERS)
    for col, _ in enumerate(HEADERS, 1):
        c = ws.cell(1, col)
        c.font      = Font(bold=True, color="FFFFFF", size=12)
        c.fill      = PatternFill("solid", fgColor=HDR_BG)
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22
    for col, w in zip("ABCDEF", [38, 28, 42, 38, 42, 42]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    wb.save(OUTPUT)
    print(f"[Excel] Created {OUTPUT}")

def append_lead_to_excel(lead):
    wb = openpyxl.load_workbook(OUTPUT)
    ws = wb.active
    ws.append([
        lead.get("name", ""),
        lead.get("area", ""),
        lead.get("website", ""),
        lead.get("email", ""),
        lead.get("instagram", ""),
        lead.get("facebook", ""),
    ])
    r = ws.max_row
    bg = EMAIL_BG if lead.get("email") else ROW_BG[(r - 2) % 2]
    for c in range(1, 7):
        cell = ws.cell(r, c)
        cell.fill      = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(vertical="center", wrap_text=False)
    wb.save(OUTPUT)

# ── Email / social extraction ─────────────────────────────────────────────────
EMAIL_RE = re.compile(r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}')

def clean_emails(raw_list):
    out = []
    seen = set()
    for e in raw_list:
        e = e.lower().strip().rstrip('.')
        if e in seen:
            continue
        seen.add(e)
        domain = e.split('@')[-1]
        if domain in BAD_EMAIL_DOMAINS:
            continue
        if any(e.startswith(p) for p in BAD_EMAIL_PREFIXES):
            continue
        if len(e) < 6 or '.' not in domain:
            continue
        out.append(e)
    return out

def extract_emails_from_html(html):
    mailto = [m.group(1).split('?')[0] for m in re.finditer(r'mailto:([^\s"\'<>]+)', html)]
    plain  = EMAIL_RE.findall(html)
    return clean_emails(mailto + plain)

def extract_socials(html):
    fb_url, ig_url = None, None
    for m in re.finditer(r'href=["\']([^"\']*facebook\.com/(?!sharer|share|dialog|tr\?|plugins)[^"\']+)["\']', html):
        fb_url = m.group(1).split('?')[0].rstrip('/'); break
    for m in re.finditer(r'href=["\']([^"\']*instagram\.com/(?!p/|reel/)[^"\']+)["\']', html):
        ig_url = m.group(1).split('?')[0].rstrip('/'); break
    return fb_url, ig_url

# ── Google Maps scraping ──────────────────────────────────────────────────────
def scroll_maps_panel(page):
    try:
        panel = page.locator('[role="feed"]').first
        for _ in range(14):
            panel.evaluate("el => el.scrollBy(0, 700)")
            sleep(0.7, 1.3)
    except Exception:
        for _ in range(10):
            try:
                page.keyboard.press("End")
                sleep(0.6, 1.0)
            except Exception:
                break

def scrape_maps_query(page, query):
    url = "https://www.google.com/maps/search/" + query.replace(" ", "+")
    print(f"\n[{ts()}] Maps search: {query}")
    try:
        page.goto(url, timeout=30000, wait_until="domcontentloaded")
        sleep(2.5, 4.0)

        for sel in ['button[aria-label*="Reject"]', 'button[aria-label*="Accept"]',
                    'form[action*="consent"] button']:
            try:
                btn = page.locator(sel).first
                if btn.is_visible(timeout=2000):
                    btn.click()
                    sleep(0.5, 1.0)
                    break
            except Exception:
                pass

        scroll_maps_panel(page)
        sleep(1.0, 2.0)

        places = []
        cards = page.locator('a.hfpxzc').all()
        if not cards:
            cards = page.locator('a[href*="/maps/place/"]').all()

        for card in cards:
            try:
                name = (card.get_attribute('aria-label') or '').strip()
                href = card.get_attribute('href') or ''
                if not name or '/maps/place/' not in href:
                    continue
                if href.startswith('//'):
                    href = 'https:' + href
                elif href.startswith('/'):
                    href = 'https://www.google.com' + href
                places.append({'name': name, 'maps_url': href})
            except Exception:
                continue

        print(f"  -> {len(places)} places found")
        return places

    except Exception as e:
        print(f"  [!] Query failed: {e}")
        return []

# ── Place detail → website ────────────────────────────────────────────────────
def get_website(page, maps_url):
    website = None
    area    = None
    try:
        page.goto(maps_url, timeout=25000, wait_until="domcontentloaded")
        sleep(1.5, 2.5)

        link = page.locator('a[data-item-id="authority"]').first
        if link.count():
            href = link.get_attribute('href') or ''
            if href.startswith('http'):
                website = href.split('?')[0]

        if not website:
            for a in page.locator('a[href^="http"]').all():
                href = a.get_attribute('href') or ''
                if href.startswith('http') and 'google' not in href and 'maps' not in href:
                    website = href.split('?')[0]
                    break

        try:
            addr_btn = page.locator('button[data-item-id="address"]').first
            if addr_btn.count():
                addr_text = addr_btn.get_attribute('aria-label') or addr_btn.inner_text() or ''
                addr_text = re.sub(r'^Address:\s*', '', addr_text, flags=re.IGNORECASE).strip()
                parts = [p.strip() for p in addr_text.split(',')]
                for p in reversed(parts[1:]):
                    if re.match(r'^[A-Z]{2}\s+\d', p):
                        continue
                    if re.match(r'^\d{4,5}', p):
                        continue
                    if re.match(r'^[A-Z]{2}$', p):
                        continue
                    if p.lower() in ('united states', 'usa', 'us', 'canada'):
                        continue
                    area = p
                    break
        except Exception:
            pass

    except Exception:
        pass
    return website, area

# ── Website scraping ──────────────────────────────────────────────────────────
def collect_html(page, base_url):
    html_parts   = []
    contact_urls = []

    try:
        page.goto(base_url, timeout=15000, wait_until="domcontentloaded")
        sleep(1.0, 2.0)
        html_parts.append(page.content())

        for a in page.locator('a[href]').all()[:80]:
            try:
                href = a.get_attribute('href') or ''
                text = (a.inner_text() or '').lower().strip()
                if not href.startswith('http'):
                    href = base_url.rstrip('/') + '/' + href.lstrip('/')
                if href in contact_urls or href == base_url:
                    continue
                if any(k in href.lower() or k in text
                       for k in ['contact', 'about', 'reach', 'connect', 'location']):
                    contact_urls.append(href)
            except Exception:
                continue

    except Exception:
        pass

    for url in contact_urls[:1]:
        try:
            page.goto(url, timeout=12000, wait_until="domcontentloaded")
            sleep(0.8, 1.5)
            html_parts.append(page.content())
        except Exception:
            continue

    return "\n".join(html_parts)

def scrape_website(page, website_url):
    try:
        html = collect_html(page, website_url)
        emails = extract_emails_from_html(html)
        fb, ig = extract_socials(html)
        return {
            'email':     emails[0] if emails else None,
            'facebook':  fb,
            'instagram': ig,
        }
    except Exception:
        return {'email': None, 'facebook': None, 'instagram': None}

# ── Facebook About scraping ───────────────────────────────────────────────────
def scrape_fb_email(page, fb_url):
    try:
        about = fb_url.rstrip('/') + '/about'
        page.goto(about, timeout=20000, wait_until="domcontentloaded")
        sleep(2.0, 3.5)
        html = page.content()
        if html.count('@') == 0 and 'email' not in html.lower():
            return None
        emails = extract_emails_from_html(html)
        return emails[0] if emails else None
    except Exception:
        return None

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--reset', action='store_true', help='Clear progress and restart')
    args = parser.parse_args()

    if args.reset:
        for f in [PROGRESS, PROGRESS + ".tmp"]:
            Path(f).unlink(missing_ok=True)
        print("[Reset] Progress cleared.")

    seen_names, seen_emails = load_seen_db()

    existing_names = load_existing_names()
    existing_count = count_existing_rows()
    print(f"[Dedup] {existing_count} leads already in {OUTPUT}")

    existing_names.update(seen_names)
    existing_emails = set(seen_emails)

    prog   = load_progress()
    leads  = prog.get("leads", [])
    done   = set(prog.get("done_urls", []))
    places = prog.get("places", [])
    q_idx  = prog.get("query_idx", 0)

    for lead in leads:
        existing_names.add(lead['name'].lower().strip())
        if lead.get('email'):
            existing_emails.add(lead['email'].lower().strip())

    if not Path(OUTPUT).exists():
        init_excel()
        for lead in leads:
            append_lead_to_excel(lead)

    print(f"[Start] {len(leads)} new leads done this run | {len(places)} places collected | query #{q_idx}")

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=HEADLESS,
            args=["--start-maximized", "--disable-blink-features=AutomationControlled"],
        )
        ctx = browser.new_context(
            viewport={"width": 1280, "height": 800},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
        )
        page = ctx.new_page()

        # ── Phase 1: Collect places from Google Maps ─────────────────────────
        print("\n=== Phase 1: Collecting places from Google Maps ===")
        while len(places) < TARGET + 150 and q_idx < len(SEARCH_QUERIES):
            new_places = scrape_maps_query(page, SEARCH_QUERIES[q_idx])
            existing_place_names = {p['name'].lower() for p in places}
            added = 0
            for p in new_places:
                if p['name'].lower() not in existing_place_names:
                    if is_alcohol_only(p['name']):
                        print(f"  [skip alcohol] {p['name']}")
                        continue
                    places.append(p)
                    existing_place_names.add(p['name'].lower())
                    added += 1
            print(f"  +{added} new unique | total {len(places)}")
            q_idx += 1
            prog.update({"query_idx": q_idx, "places": places})
            save_progress(prog)
            sleep(2.0, 4.0)

        print(f"\n[Phase 1 done] {len(places)} unique places collected")
        if len(places) < 50:
            print("[!] Too few places - Google may be blocking. Try again later or solve CAPTCHA manually.")

        # ── Phase 2: Process each place ──────────────────────────────────────
        print("\n=== Phase 2: Scraping websites + emails ===")
        for place in places:
            if len(leads) >= TARGET:
                break

            maps_url = place['maps_url']
            name     = place['name']

            if maps_url in done:
                continue

            if name.lower().strip() in existing_names:
                print(f"  [skip seen name] {name}")
                done.add(maps_url)
                continue

            print(f"\n[{ts()}] [{len(leads)+1}/{TARGET}] {name}")

            website, area = get_website(page, maps_url)
            print(f"  Website : {website or '-'}")
            print(f"  Area    : {area or '-'}")

            lead = {
                'name':      name,
                'area':      area or '',
                'website':   website or '',
                'email':     '',
                'instagram': '',
                'facebook':  '',
            }

            if website:
                data = scrape_website(page, website)
                lead['email']     = data.get('email') or ''
                lead['facebook']  = data.get('facebook') or ''
                lead['instagram'] = data.get('instagram') or ''
                print(f"  Email   : {lead['email'] or '-'}")
                print(f"  FB      : {lead['facebook'] or '-'}")
                print(f"  IG      : {lead['instagram'] or '-'}")

                if not lead['email'] and lead['facebook']:
                    print(f"  -> Checking FB About for email...")
                    fb_email = scrape_fb_email(page, lead['facebook'])
                    if fb_email:
                        lead['email'] = fb_email
                        print(f"  -> FB email found: {fb_email}")

            if lead['email'] and lead['email'].lower().strip() in existing_emails:
                print(f"  [skip dup email] {name} — {lead['email']} already recorded")
                done.add(maps_url)
                continue

            leads.append(lead)
            existing_names.add(name.lower().strip())
            if lead['email']:
                existing_emails.add(lead['email'].lower().strip())
            done.add(maps_url)
            prog.update({"leads": leads, "done_urls": list(done)})
            save_progress(prog)
            append_lead_to_excel(lead)
            _save_seen_db(existing_names, existing_emails)

            sleep(1.5, 3.0)

        browser.close()

    with_email   = sum(1 for l in leads if l.get('email'))
    with_fb      = sum(1 for l in leads if l.get('facebook'))
    with_ig      = sum(1 for l in leads if l.get('instagram'))
    with_website = sum(1 for l in leads if l.get('website'))

    print(f"""
=== Done ===
New leads added : {len(leads)}
With website    : {with_website}
With email      : {with_email}
With Facebook   : {with_fb}
With Instagram  : {with_ig}
Total in file   : {existing_count + len(leads)}
Saved to        : {OUTPUT}
""")

if __name__ == "__main__":
    main()

import sys, io, json, urllib.request, urllib.error, os, time, re
from dotenv import load_dotenv

# Fix: force UTF-8 stdout so Chinese characters don't crash on Windows cp1252
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

import openpyxl

load_dotenv()
API_KEY = os.getenv("GEMINI_API_KEY", "")
if not API_KEY:
    print("ERROR: GEMINI_API_KEY not found in .env")
    sys.exit(1)

MODELS = [
    "gemini-2.5-flash",
    "gemini-2.5-flash-lite",
]

# Usage: py generate_emails.py <input_xlsx>
# Output file is auto-named: <input_basename>_emails.xlsx in the same folder
if len(sys.argv) < 2:
    print("Usage: py generate_emails.py <path/to/qualified.xlsx>")
    sys.exit(1)

INPUT_FILE = sys.argv[1]
base, ext = os.path.splitext(INPUT_FILE)
OUTPUT_FILE = base + "_emails" + ext

# Extract area name from filename: "Cherry Creek North Denver leads.xlsx" → "Cherry Creek North, Denver"
_filename_base = os.path.basename(base)
AREA = re.sub(r'[_\-]emails$', '', _filename_base, flags=re.IGNORECASE)
AREA = re.sub(r'\s*leads\s*$', '', AREA, flags=re.IGNORECASE).strip()
print(f"Detected area: {AREA}")


def call_gemini(biz_name, pain_point, owner_name=None, menu_items=None):
    subject_name = owner_name.split()[0] if owner_name else biz_name
    subject = f"{subject_name} x Y-Studio"

    context_lines = []
    if pain_point:
        context_lines.append(f"Observed pain point: {pain_point}")
    if menu_items:
        context_lines.append(f"Known menu items: {menu_items}")
    context_block = ("\n" + "\n".join(context_lines)) if context_lines else ""

    prompt = f"""You are an expert Cold Email Copywriter specializing in B2B outreach and high-converting marketing for the food & beverage, lifestyle, and retail industries.

Your goal is to write a highly compelling, personalized, and short Cold Email to the owner of "{biz_name}".

The business is located in: {AREA}{context_block}

The core psychology of the email is to leverage "competitor anxiety" and offer a "visual upgrade solution" to outshine their local competitor, backed by proven success.

Y-Studio specializes in creating high-end food poster design — visually striking posters that make food look irresistible on Instagram and in-store displays. IMPORTANT: Never use the words "photography" or "photo shoot" — Y-Studio is a poster design studio, not a photography studio. Always say "food poster", "poster design", or "poster content".

Write exactly 3 tight paragraphs. No subject line. No bullet points. No labels.

Para 1 — Salutation + Hook: Start with "Hi {biz_name}," (always use the exact business name, never a placeholder). Then name a specific, plausible competitor from {AREA} (invent a realistic-sounding name that fits the neighbourhood). Describe vividly how their consistent food poster content on Instagram racks up engagement and pulls walk-ins — make the reader feel the gap as if they are scrolling past those posts right now.

Para 2 — Flattery and Pivot: Open with "But what I noticed is..." or "But here's the thing —". Give a specific, believable compliment tied to their actual concept, neighbourhood identity, or loyal customer base. If menu items are provided, naturally reference 1–2 specific dishes or drinks by name to show you actually know their food — NOT generic phrases like "your food is incredible". Then say it is a shame a competitor with weaker food is winning purely through stronger poster content.

Para 3 — Value Prop + CTA: Say you have already spotted 2 food poster design ideas for {biz_name} that could flip this. If menu items are known, hint that these ideas are built around their specific dishes. Drop one concrete number from a past client (e.g. "3x reach in 6 weeks"). End with one zero-pressure sentence asking if they want to see the 2 ideas and the case study.

Tone: Professional, confident, helpful, peer-to-peer — sounds like a real human wrote it, not a template. No corporate jargon.
Length: Under 120 words. Hard limit. Cut ruthlessly if needed.

Return ONLY the email body. No subject line, no labels, no extra text."""

    payload = json.dumps({
        "contents": [{"parts": [{"text": prompt}]}]
    }).encode("utf-8")

    for model in MODELS:
        endpoint = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={API_KEY}"
        for attempt in range(3):
            try:
                req = urllib.request.Request(endpoint, data=payload, headers={"Content-Type": "application/json"})
                with urllib.request.urlopen(req, timeout=60) as resp:
                    result = json.loads(resp.read())
                text = result["candidates"][0]["content"]["parts"][0]["text"].strip()
                return subject, text
            except urllib.error.HTTPError as e:
                wait = 5 * (attempt + 1)
                print(f"  [{model}] attempt {attempt+1} failed ({e.code}) — retry in {wait}s")
                time.sleep(wait)
            except Exception as e:
                wait = 5 * (attempt + 1)
                print(f"  [{model}] attempt {attempt+1} failed ({type(e).__name__}) — retry in {wait}s")
                time.sleep(wait)
        print(f"  [{model}] all attempts failed, trying next model...")

    raise RuntimeError("All models exhausted")


def clean_body(text):
    # Remove markdown bold/italic markers and excessive whitespace
    cleaned = re.sub(r"[*_`#]", "", text)
    cleaned = re.sub(r" {2,}", " ", cleaned)
    return cleaned.strip()


# Resume: load OUTPUT_FILE if exists, else start from INPUT_FILE
import shutil
if not os.path.exists(OUTPUT_FILE):
    shutil.copy2(INPUT_FILE, OUTPUT_FILE)
    print(f"Created output file: {OUTPUT_FILE}")
else:
    print(f"Resuming from existing: {OUTPUT_FILE}")

wb = openpyxl.load_workbook(OUTPUT_FILE)
ws = wb.active
headers = [cell.value for cell in ws[1]]

for col_name in ["Subject", "Cold Email"]:
    if col_name not in headers:
        headers.append(col_name)
        ws.cell(row=1, column=len(headers), value=col_name)

subject_col = headers.index("Subject") + 1
email_col = headers.index("Cold Email") + 1
name_col_key = next((k for k in ["Restaurant Name", "Name", "name"] if k in headers), headers[0])
name_col = headers.index(name_col_key) + 1
# All columns optional except Name
pain_col  = headers.index("Pain Point") + 1 if "Pain Point" in headers else None
owner_col = headers.index("Owner") + 1 if "Owner" in headers else None
menu_col  = headers.index("Menu Items") + 1 if "Menu Items" in headers else None

total = ws.max_row - 1
done = 0
skipped = 0

for row_idx in range(2, ws.max_row + 1):
    biz_name   = ws.cell(row=row_idx, column=name_col).value
    pain_point = ws.cell(row=row_idx, column=pain_col).value if pain_col else ""
    if not biz_name:
        continue

    # Skip if Email already filled (resume checkpoint)
    existing_email = ws.cell(row=row_idx, column=email_col).value or ""
    if existing_email and existing_email.startswith("Hi "):
        skipped += 1
        print(f"[{row_idx-1}/{total}] SKIP (already done): {biz_name}")
        continue

    owner_name  = ws.cell(row=row_idx, column=owner_col).value if owner_col else None
    menu_items  = ws.cell(row=row_idx, column=menu_col).value if menu_col else None

    print(f"[{row_idx-1}/{total}] Generating: {biz_name} ...")
    try:
        subject, email_body = call_gemini(biz_name, pain_point, owner_name, menu_items)
    except Exception as e:
        print(f"  ERROR: {e} — skipping, will retry on next run")
        continue

    email_body = clean_body(email_body)

    ws.cell(row=row_idx, column=subject_col, value=subject)
    ws.cell(row=row_idx, column=email_col, value=email_body)

    wb.save(OUTPUT_FILE)  # save immediately after each row
    done += 1

    print(f"  Subject: {subject}")
    print(f"  Email: {email_body[:100]}...")
    time.sleep(3)  # avoid rate-limiting

print(f"\nDone. {done} generated, {skipped} skipped. Output: {OUTPUT_FILE}")

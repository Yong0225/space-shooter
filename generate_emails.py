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


def call_gemini(biz_name, food_post=None, menu_items=None):
    subject = f"{biz_name} x Y-Studio"

    if food_post:
        food_hint = f"The food item from their recent post: {food_post}"
    elif menu_items:
        food_hint = f"Their menu includes: {menu_items}. Pick the single most visually appealing / photogenic item from this list."
    else:
        food_hint = "Pick any one common, visually appealing food item (e.g. burger, ramen, fried chicken, tacos, pasta, boba, waffle, etc.)."

    prompt = f"""Write a short, casual Instagram DM / cold outreach message for Y-Studio, a food poster design studio.

Target business: {biz_name}
{food_hint}

Follow this structure exactly (2–3 sentences, under 45 words total):
Sentence 1: "Hey! [Casual opener — say you were scrolling and came across / just saw / spotted their recent [food item] post.]"
Sentence 2: "[One-line compliment on how the food looks — e.g. 'the food honestly already looks really good', 'the food presentation has a lot of potential', 'it honestly looks super solid'.]"
Sentence 3: "Had [a bit of / some] free time so I [made / turned it into / put together] a [promo-style concept / poster-style visual / scroll-stopping promo visual] using your [photo / original photo / shot]. Mind if I send it over?"

Rules:
- Vary the exact wording each time — never copy sentence-for-sentence from the examples.
- Keep it conversational and human. No hashtags, no emojis, no salesy language.
- Do NOT mention Y-Studio by name in the message body.
- Return ONLY the message. No subject line, no labels, no extra text."""

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
email_col   = headers.index("Cold Email") + 1
name_col_key = next((k for k in ["Restaurant Name", "Name", "name"] if k in headers), headers[0])
name_col    = headers.index(name_col_key) + 1
food_col    = headers.index("Food Post") + 1 if "Food Post" in headers else None
menu_col    = headers.index("Menu Items") + 1 if "Menu Items" in headers else None

total   = ws.max_row - 1
done    = 0
skipped = 0

for row_idx in range(2, ws.max_row + 1):
    biz_name = ws.cell(row=row_idx, column=name_col).value
    if not biz_name:
        continue

    # Skip if email already filled (resume checkpoint)
    existing_email = ws.cell(row=row_idx, column=email_col).value or ""
    if existing_email and existing_email.lower().startswith("hey"):
        skipped += 1
        print(f"[{row_idx-1}/{total}] SKIP (already done): {biz_name}")
        continue

    food_post  = ws.cell(row=row_idx, column=food_col).value if food_col else None
    menu_items = ws.cell(row=row_idx, column=menu_col).value if menu_col else None

    print(f"[{row_idx-1}/{total}] Generating: {biz_name} ...")
    try:
        subject, email_body = call_gemini(biz_name, food_post, menu_items)
    except Exception as e:
        print(f"  ERROR: {e} — skipping, will retry on next run")
        continue

    email_body = clean_body(email_body)

    ws.cell(row=row_idx, column=subject_col, value=subject)
    ws.cell(row=row_idx, column=email_col,   value=email_body)

    wb.save(OUTPUT_FILE)
    done += 1

    print(f"  Subject: {subject}")
    print(f"  Email: {email_body[:120]}...")
    time.sleep(3)

print(f"\nDone. {done} generated, {skipped} skipped. Output: {OUTPUT_FILE}")

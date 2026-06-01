import sys, io, json, urllib.request, urllib.error, os, time, re, random
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


def call_gemini(biz_name, food_post=None, menu_items=None, row_seed=None):
    subject = biz_name

    prompt = """You are an elite cold email copywriter trained in Jeremy Miner's NEPQ methodology.

Generate a cold email for a restaurant owner.

STRICT 3-LINE STRUCTURE — follow exactly:
Line 1: "Hey,"
Line 2: One sentence — saw/came across/was checking out + one of: "your page" / "your social page" / "one of your posts" + optional casual time reference (earlier, today, just now).
Line 3: Curiosity opener + one soft question. Curiosity openers: "Just curious —" / "Quick question —" / "Out of curiosity," / "Just wondering —". Optionally end with "Just wondering." as a soft closer on a new line.

The question must be a close variation of one of these 3 styles — rotate between them:
Style A: "is that one of the menu items you'd like more people noticing online?"
Style B: "do you feel your social posts are getting the attention you'd hoped for lately?"
Style C: "do customers order it as often as you'd expect from how good it looks?"

Rules:
* Maximum 35 words total.
* Sound completely human.
* Never sell, pitch, mention services, design, agency, pricing, or marketing.
* Do NOT use the business name.
* Do NOT use placeholders.
* Generate a unique variation each time.

The goal is ONLY to start a conversation.

Good examples (follow this tone and structure exactly):

"Hey,

Saw one of your posts earlier.

Just curious — is that one of the menu items you'd like more people noticing online?"

---

"Hey,

Was checking out your social page.

Quick question — do you feel your social posts are getting the attention you'd hoped for lately?"

---

"Hey,

Came across one of your posts today.

Just wondering — do customers order it as often as you'd expect from how good it looks?"

---

"Hey,

Saw one of your posts just now.

Out of curiosity — is that one of the dishes you'd love more people discovering online? Just wondering."

Return ONLY the email message, nothing else."""

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
    # Replace dash used as punctuation separator with comma (space-dash-space, or em/en dash)
    cleaned = re.sub(r" - ", ", ", cleaned)
    cleaned = re.sub(r"\s*[–—]\s*", ", ", cleaned)
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
        subject, email_body = call_gemini(biz_name, food_post, menu_items, row_seed=row_idx)
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

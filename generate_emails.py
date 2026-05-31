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


# ── Phrase banks: one slot drawn per email to force wording variation ──────────
_OPENERS = [
    "Just saw one of your {food} posts while scrolling",
    "Was scrolling earlier and came across one of your {food} posts",
    "Came across one of your {food} posts while scrolling",
    "Spotted one of your {food} posts while I was scrolling",
    "Just stumbled across one of your {food} posts",
    "One of your {food} posts just popped up while I was scrolling",
    "Was just scrolling and noticed one of your {food} posts",
    "Just came across one of your {food} posts on my feed",
    "Saw one of your {food} posts while scrolling",
    "Was scrolling and one of your {food} posts caught my eye",
]

_COMPLIMENTS = [
    "the food honestly already looks really good",
    "the food presentation honestly has a lot of potential",
    "it honestly looks super solid",
    "honestly the food already looks amazing",
    "the {food} honestly looks really appetizing",
    "the food already looks really well-done",
    "it honestly already looks really appetizing",
    "the food honestly looks really promising",
    "ngl the food already looks pretty fire",
    "the food genuinely looks great as is",
]

_FREETIME = [
    "Had a bit of free time so I",
    "Had some free time so I",
    "Was messing around and",
    "Had a spare moment so I",
    "Just had some downtime so I",
]

_ACTIONS = [
    "made a quick promo-style concept using your photo",
    "turned it into a poster-style visual just for fun",
    "put together a scroll-stopping promo visual using your original photo",
    "put together a quick poster concept around your shot",
    "made a promo-style design using your photo",
    "turned your photo into a clean poster-style visual",
    "put together a quick food poster concept from your shot",
    "turned it into a promo design using your photo",
    "made a scroll-stopping version of it using your original shot",
    "put together a visual concept using your photo",
]

_CTAS = [
    "Mind if I send it over?",
    "Mind if I share it with you?",
    "Mind if I pass it along?",
    "Would it be okay if I sent it over?",
    "Happy to send it over if you'd like to take a look!",
]


def _pick(lst, food=""):
    item = random.choice(lst)
    return item.replace("{food}", food)


def call_gemini(biz_name, food_post=None, menu_items=None, row_seed=None):
    subject = f"{biz_name} x Y-Studio"

    # Seed per-row so retries reproduce the same combo but each row differs
    rng = random.Random(row_seed)

    if food_post:
        food_hint = f"The food item from their recent post: {food_post}"
        food_label = food_post
    elif menu_items:
        food_hint = f"Their menu includes: {menu_items}. Pick the single most visually appealing / photogenic item."
        food_label = "{food}"  # Gemini will resolve this
    else:
        food_hint = "Use a generic word like 'food', 'dish', or 'drinks' — do NOT name a specific food item."
        food_label = "food"

    opener    = rng.choice(_OPENERS).replace("{food}", food_label)
    compliment = rng.choice(_COMPLIMENTS).replace("{food}", food_label)
    freetime  = rng.choice(_FREETIME)
    action    = rng.choice(_ACTIONS)
    cta       = rng.choice(_CTAS)

    prompt = f"""Write a short, casual Instagram DM for a food poster design studio.

Target business: {biz_name}
{food_hint}

Build the message around these exact phrase seeds — rephrase them naturally, do NOT copy word-for-word, but keep the core idea of each:
• Opener seed: "{opener}"
• Compliment seed: "{compliment}"
• Action seed: "{freetime} {action}"
• CTA: "{cta}"

Tone and reading level: Write at a grade school level. Use short, simple words and short sentences. Avoid any fancy, formal, or complex vocabulary. Sound like a real person texting a friend, not a professional writer.

Output format: 2–3 sentences, under 45 words. Start with "Hey!". No hashtags, no emojis, no salesy language. Do NOT name the studio. Do NOT mention specific food items unless provided above. Do NOT use words like "recent", "latest", or "newest" — say "one of your posts" instead. Return ONLY the message."""

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

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


def call_gemini(biz_name, pain_point, owner_name=None):
    subject_name = owner_name.split()[0] if owner_name else biz_name
    subject = f"{subject_name} x Y-Studio"

    prompt = f"""You are an expert Cold Email Copywriter specializing in B2B outreach and high-converting marketing for the food & beverage, lifestyle, and retail industries.

Your goal is to write a highly compelling, personalized, and short Cold Email to the owner of "{biz_name}".

The business is located in: {AREA}
The observed pain point about this business is: {pain_point}

The core psychology of the email is to leverage "competitor anxiety" and offer a "visual upgrade solution" to outshine their local competitor, backed by proven success.

Y-Studio specializes in creating high-end food poster design — visually striking posters that make food look irresistible on Instagram and in-store displays. IMPORTANT: Never use the words "photography" or "photo shoot" — Y-Studio is a poster design studio, not a photography studio. Always say "food poster", "poster design", or "poster content".

Follow this strict structure and write it as ONE flowing paragraph or 2–3 short paragraphs — not as separate bullet-point blocks:

1. Do NOT include a subject line — the subject is already set. Start directly with the salutation "Hi {biz_name}," (or use the owner first name if it feels natural).
2. Hook: Mention that a specific, plausible competitor from the {AREA} area (invent a realistic-sounding name that fits that neighbourhood) has been posting strong food poster content on social media and pulling in walk-ins because of it.
3. Flattery & Pivot: Flow naturally from the competitor result into a soft compliment — use a transition like "But what I noticed is..." or "But honestly..." to say that {biz_name}'s food is actually better, and with the right poster content it can absolutely outshine them.
4. Value Proposition + Social Proof: Naturally lead into the fact that you've spotted 2 specific food poster design ideas for {biz_name} that could flip the game. Then weave in a brief success story (e.g. "We recently helped a local cafe in a similar spot grow their Instagram engagement 3x in 6 weeks with a new food poster series") — make it feel like a natural aside, not a sudden announcement.
5. Low-friction CTA: One easy closing sentence asking if they'd like to see the 2 custom poster ideas and the quick case study (e.g. "I've put these 2 ideas together along with the case study — want me to send them over? Just reply 'yes'.")

Tone: Professional, confident, helpful, peer-to-peer — sounds like a real human wrote it, not a template. No corporate jargon. The whole email should read as one natural, connected thought, not a list of blocks.
Length: Under 100 words. This is a strict hard limit. Count every word before finalizing. If over 100, cut ruthlessly — shorten sentences, remove filler words. Do not go over.

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

for col_name in ["Subject", "Email"]:
    if col_name not in headers:
        headers.append(col_name)
        ws.cell(row=1, column=len(headers), value=col_name)

subject_col = headers.index("Subject") + 1
email_col = headers.index("Email") + 1
name_col_key = next((k for k in ["Restaurant Name", "Name", "name"] if k in headers), headers[0])
name_col = headers.index(name_col_key) + 1
pain_col = headers.index("Pain Point") + 1
# Owner column is optional — used for subject line if present
owner_col = headers.index("Owner") + 1 if "Owner" in headers else None

total = ws.max_row - 1
done = 0
skipped = 0

for row_idx in range(2, ws.max_row + 1):
    biz_name = ws.cell(row=row_idx, column=name_col).value
    pain_point = ws.cell(row=row_idx, column=pain_col).value
    if not biz_name or not pain_point:
        continue

    # Skip if Email already filled (resume checkpoint)
    existing_email = ws.cell(row=row_idx, column=email_col).value or ""
    if existing_email and existing_email.startswith("Hi "):
        skipped += 1
        print(f"[{row_idx-1}/{total}] SKIP (already done): {biz_name}")
        continue

    owner_name = ws.cell(row=row_idx, column=owner_col).value if owner_col else None

    print(f"[{row_idx-1}/{total}] Generating: {biz_name} ...")
    try:
        subject, email_body = call_gemini(biz_name, pain_point, owner_name)
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

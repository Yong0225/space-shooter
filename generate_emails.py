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


def call_gemini(biz_name, pain_point, owner_name=None):
    subject_name = owner_name.split()[0] if owner_name else biz_name
    subject = f"{subject_name} x Y-Studio"

    prompt = f"""You are an expert Cold Email Copywriter specializing in B2B outreach and high-converting marketing for the food & beverage, lifestyle, and retail industries.

Your goal is to write a highly compelling, personalized, and short Cold Email to the owner of "{biz_name}".

The observed pain point about this business is: {pain_point}

The core psychology of the email is to leverage "competitor anxiety" and offer a "visual upgrade solution" to outshine their local competitor, backed by proven success.

Y-Studio specializes in creating high-end food poster design — visually striking posters that make food look irresistible on Instagram and in-store displays. IMPORTANT: Never use the words "photography" or "photo shoot" — Y-Studio is a poster design studio, not a photography studio. Always say "food poster", "poster design", or "poster content".

Follow this strict structure:

1. Do NOT include a subject line — the subject is already set. Start directly with the salutation "Hi {biz_name}," (or use the owner first name if it feels natural).
2. Hook: Mention that their local competitor (invent a plausible-sounding nearby competitor of the same business type) is posting stunning food posters on social media and getting strong results (e.g. growing following fast, pulling walk-ins, filling seats).
3. Flattery & Pivot: Immediately after the competitor result, pivot to compliment the prospect — say that their food is actually better than the competitor's and absolutely has what it takes to outshine them, but right now the competitor is winning the audience purely because of stronger poster content.
4. Value Proposition + Social Proof: State that you have identified 2 specific food poster design ideas that would make their dishes look irresistible online and pull more attention than the competitor. Support it with a brief success story — e.g. "We recently helped a local cafe grow their Instagram engagement 3x in 6 weeks with a new food poster series" (keep it short and believable).
5. Low-friction CTA: One sentence only — ask if they want to see the 2 custom poster ideas and the case study. Example: "I've put together these 2 custom poster ideas along with the quick case study of how we did it. Open to checking them out? Just reply 'yes' and I'll send them over."

Tone: Professional, confident, helpful, peer-to-peer — sounds like a real human, not a template. No corporate jargon.
Length: Under 120 words. Hard limit — count every word before finalizing.

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

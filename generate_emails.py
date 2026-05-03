import sys, io, json, urllib.request, os
from dotenv import load_dotenv

# Fix: force UTF-8 stdout so Chinese characters don't crash on Windows cp1252
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

import openpyxl

load_dotenv()
API_KEY = os.getenv("GEMINI_API_KEY", "")
if not API_KEY:
    print("ERROR: GEMINI_API_KEY not found in .env")
    sys.exit(1)
ENDPOINT = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={API_KEY}"

# Usage: py generate_emails.py <input_xlsx>
# Output file is auto-named: <input_basename>_emails.xlsx in the same folder
if len(sys.argv) < 2:
    print("Usage: py generate_emails.py <path/to/qualified.xlsx>")
    sys.exit(1)

INPUT_FILE = sys.argv[1]
base, ext = os.path.splitext(INPUT_FILE)
OUTPUT_FILE = base + "_emails" + ext


def call_gemini(business_name, pain_point):
    prompt = f"""You are an expert cold email copywriter for a social media marketing agency called Y-Studio.
Write 2 different cold EMAIL versions for a restaurant/cafe called "{business_name}".

The observed pain point is: {pain_point}

Use the PEC formula for both versions:
- P (Pain/Fear): Open with "I noticed..." referencing the pain point, then highlight the fear or desire behind it — what they are losing or missing out on.
- E (Evidence): Mention that you designed a demo specifically for {business_name} and what result it could drive (more walk-ins, more sales, more leads, or stronger engagement).
- C (CTA): End with a soft, one-line ask — invite them to see the demo.

Rules:
- Format: cold EMAIL body only (no subject line, no sign-off)
- Salutation: start with "Hi {business_name},"
- Tone: casual, direct, human — like a real person reaching out, not a pitch deck
- Length: 60–90 words each
- No hashtags, no emojis, no salesy buzzwords
- Never use words like "free", "guaranteed", "limited offer", "marketing agency", "social media services"
- The 2 versions must feel noticeably different — vary the fear angle, the evidence framing, or the CTA wording
- Do not use the word "just" or "I wanted to"

Return ONLY the 2 email bodies separated by "---", no labels, no numbering, no extra text."""

    payload = json.dumps({
        "contents": [{"parts": [{"text": prompt}]}]
    }).encode("utf-8")

    req = urllib.request.Request(ENDPOINT, data=payload, headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req) as resp:
        result = json.loads(resp.read())

    text = result["candidates"][0]["content"]["parts"][0]["text"].strip()
    parts = text.split("---")
    v1 = parts[0].strip() if len(parts) >= 1 else ""
    v2 = parts[1].strip() if len(parts) >= 2 else ""
    return v1, v2


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

for col_name in ["Subject", "Email v1", "Email v2"]:
    if col_name not in headers:
        headers.append(col_name)
        ws.cell(row=1, column=len(headers), value=col_name)

subject_col = headers.index("Subject") + 1
v1_col = headers.index("Email v1") + 1
v2_col = headers.index("Email v2") + 1
name_col_key = "Restaurant Name" if "Restaurant Name" in headers else "name"
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

    # Skip if Email v1 already filled (resume checkpoint)
    if ws.cell(row=row_idx, column=v1_col).value:
        skipped += 1
        print(f"[{row_idx-1}/{total}] SKIP (already done): {biz_name}")
        continue

    # Subject priority: owner first name > business name
    owner = ws.cell(row=row_idx, column=owner_col).value if owner_col else None
    first_name = owner.strip().split()[0] if owner and str(owner).strip() else None
    subject = f"Hi {first_name}" if first_name else f"{biz_name} x Y-Studio"

    print(f"[{row_idx-1}/{total}] Generating: {biz_name} ...")
    v1, v2 = call_gemini(biz_name, pain_point)

    ws.cell(row=row_idx, column=subject_col, value=subject)
    ws.cell(row=row_idx, column=v1_col, value=v1)
    ws.cell(row=row_idx, column=v2_col, value=v2)

    wb.save(OUTPUT_FILE)  # save immediately after each row
    done += 1

    print(f"  Subject: {subject}")
    print(f"  v1: {v1[:80]}...")
    print(f"  v2: {v2[:80]}...")

print(f"\nDone. {done} generated, {skipped} skipped. Output: {OUTPUT_FILE}")

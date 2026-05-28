#!/usr/bin/env python3
"""
Usage: py clean_leads.py "Some leads.xlsx"
  - Removes rows with no email
  - Makes Instagram, Facebook, and Website cells clickable hyperlinks
"""

import sys
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

def clean(path_str):
    path = Path(path_str)
    if not path.exists():
        print(f"[!] File not found: {path}")
        return

    wb = openpyxl.load_workbook(path)
    ws = wb.active

    # Detect column indices from header row
    headers = [c.value for c in ws[1]]
    col = {str(h).strip().lower(): i+1 for i, h in enumerate(headers) if h}
    email_col   = col.get("email")
    ig_col      = col.get("instagram")
    fb_col      = col.get("facebook")
    web_col     = col.get("website")

    if not email_col:
        print("[!] Could not find Email column.")
        return

    HDR_BG  = "2E4057"
    ROW_BG  = ["FFFFFF", "EDF2F7"]
    EMAIL_BG = "C6F6D5"

    # Collect rows to keep (has email)
    keep = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        email = row[email_col - 1]
        if email and str(email).strip():
            keep.append(list(row))

    # Rebuild sheet: clear data rows, rewrite
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.value = None
            cell.hyperlink = None

    link_font     = Font(color="0563C1", underline="single")
    normal_font   = Font(color="000000", underline=None)

    for r_idx, row_data in enumerate(keep, start=2):
        bg = EMAIL_BG  # all kept rows have email
        for c_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(r_idx, c_idx)
            cell.value = value
            cell.fill      = PatternFill("solid", fgColor=bg)
            cell.alignment = Alignment(vertical="center", wrap_text=False)

            url = str(value).strip() if value else ""
            if url.startswith("http") and c_idx in (
                [ig_col, fb_col, web_col] if web_col else [ig_col, fb_col]
            ):
                cell.hyperlink = url
                cell.font = link_font
            else:
                cell.font = normal_font

    removed = 0
    # Count how many were removed — original data rows minus kept
    # (we can't count from ws anymore since we cleared it, so track separately)
    wb.save(path)
    print(f"[Done] {len(keep)} rows kept, saved to {path.name}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: py clean_leads.py \"filename.xlsx\"")
        sys.exit(1)
    clean(sys.argv[1])

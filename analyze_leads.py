"""
Lead Qualifier — ICP Pain Point Analyzer
Primary: Playwright screenshots + Gemini Vision (sees actual posts, poster quality, feed content)
Fallback: Google Search grounding (if page scraping fails)

Usage:
    py analyze_leads.py <input_xlsx> <output_xlsx>

Dependencies:
    py -m pip install requests openpyxl python-dotenv playwright
    py -m playwright install chromium

Environment (.env):
    GEMINI_API_KEY=<your key>
"""

import os, re, json, sys, time, base64
import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from dotenv import load_dotenv

load_dotenv()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
if not GEMINI_API_KEY:
    print("ERROR: GEMINI_API_KEY not found in .env")
    sys.exit(1)

GEMINI_URL = (
    "https://generativelanguage.googleapis.com/v1beta/models/"
    "gemini-2.5-flash:generateContent?key=" + GEMINI_API_KEY
)

RATE_LIMIT_DELAY = 3


# ─── Playwright page scraping ────────────────────────────────────────────────

def detect_platform(url: str) -> str:
    if "instagram.com" in url:
        return "instagram"
    if "facebook.com" in url or "fb.com" in url:
        return "facebook"
    return "unknown"


def _dismiss_facebook_popups(page):
    """Click Accept/Allow on Facebook's cookie consent banner."""
    selectors = [
        '[data-cookiebanner="accept_button"]',
        'div[aria-label="Allow all cookies"]',
        'button[title="Allow all cookies"]',
        'div[aria-label="Accept all"]',
        'button[aria-label="Accept all"]',
    ]
    for sel in selectors:
        try:
            loc = page.locator(sel).first
            if loc.is_visible(timeout=2000):
                loc.click()
                page.wait_for_timeout(1200)
                return True
        except Exception:
            pass
    return False


def _dismiss_instagram_popups(page):
    """
    Close Instagram's login modal.
    First tries the X button inside the dialog; if not found, presses Escape.
    """
    close_selectors = [
        'div[role="dialog"] button[aria-label="Close"]',
        'div[role="dialog"] svg[aria-label="Close"]',
        'button[aria-label="Close"]',
    ]
    for sel in close_selectors:
        try:
            loc = page.locator(sel).first
            if loc.is_visible(timeout=3000):
                loc.click()
                page.wait_for_timeout(600)
                return True
        except Exception:
            pass
    # Fallback: Escape key
    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(600)
    except Exception:
        pass
    return False


def scrape_page(url: str) -> dict:
    """
    Open the social media URL with a real (headless) browser.
    1. Wait for page to fully render (not just DOM ready)
    2. Dismiss cookie banners / login modals before screenshotting
    3. Scroll to bring the feed into view
    4. Wait for at least one post/image element to appear
    5. Capture screenshot + full page text

    Returns:
        {
            screenshot_b64: str | None,
            page_text: str,
            platform: str,
            error: str | None
        }
    """
    try:
        from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
    except ImportError:
        return {
            "screenshot_b64": None,
            "page_text": "",
            "platform": detect_platform(url),
            "error": "playwright not installed — run: py -m pip install playwright && py -m playwright install chromium",
        }

    platform = detect_platform(url)
    result = {"screenshot_b64": None, "page_text": "", "platform": platform, "error": None}

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        ctx = browser.new_context(
            viewport={"width": 1366, "height": 900},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            locale="en-US",
        )
        page = ctx.new_page()

        try:
            page.goto(url, wait_until="domcontentloaded", timeout=30000)
            # Wait for JS to render initial content
            page.wait_for_timeout(3500)

            if platform == "facebook":
                _dismiss_facebook_popups(page)
                page.wait_for_timeout(1000)

                # Scroll down past cover photo to bring timeline into view
                page.evaluate("window.scrollBy(0, 480)")
                page.wait_for_timeout(1500)

                # Wait for at least one post element
                try:
                    page.wait_for_selector(
                        '[role="article"], [data-pagelet*="Timeline"], '
                        '[data-visualcompletion="media-vc-image"]',
                        timeout=7000,
                    )
                    page.wait_for_timeout(1000)
                except Exception:
                    pass  # Take screenshot regardless

            elif platform == "instagram":
                # Wait briefly then close login modal
                page.wait_for_timeout(1500)
                _dismiss_instagram_popups(page)
                page.wait_for_timeout(800)

                # Wait for grid images to appear
                try:
                    page.wait_for_selector("article img, ._aagv img", timeout=6000)
                except Exception:
                    pass

            # Grab page text (contains follower count, post dates, like/comment counts)
            try:
                result["page_text"] = page.inner_text("body")
            except Exception:
                result["page_text"] = ""

            # Screenshot of current viewport (feed is now in view)
            screenshot_bytes = page.screenshot(type="jpeg", quality=88)
            result["screenshot_b64"] = base64.b64encode(screenshot_bytes).decode()

        except Exception as e:
            result["error"] = str(e)
        finally:
            browser.close()

    return result


# ─── Gemini API ──────────────────────────────────────────────────────────────

# Vision prompt: used when we have a real screenshot
ICP_VISION_PROMPT = """
You are a social media analyst evaluating an F&B business's social media page
for a cold email lead generation campaign.

QUALIFYING CRITERIA — lead qualifies if ANY apply:
1. Followers < 50,000
2. Posts roughly once a month or less
3. No new post for more than 1 month
4. Feed dominated by casual phone snapshots (unedited, poor lighting, random composition)
5. Feed is almost entirely Reels with NO food design posters — ONLY counts if followers < 20,000
6. Has food design posters but quality is poor (basic Canva templates, bad fonts/colors, messy layout)
7. Rarely posts food-related content (but has posted some food content before)

DISQUALIFYING CRITERIA — exclude if ANY apply:
- Followers >= 50,000
- Feed is primarily professional photography (studio lighting, sharp focus, consistent visual style)
- Never posted any food content at all
- Feed is all Reels AND followers >= 20,000
- Not an F&B business (e.g. training school, non-food business)

──────────────────────────────
BUSINESS INFO:
Name: {name}
Platform: {platform}
URL: {url}

PAGE TEXT — extracted live from the page.
Contains: follower/like counts, post timestamps, like/comment numbers.
---
{page_text}
---
──────────────────────────────

A screenshot of the page is attached.

WHAT TO LOOK FOR IN THE SCREENSHOT:
• Are posts phone snapshots (poor lighting, casual framing) or professional photos?
• Are there design posters? Do they look professionally designed or amateurish (Canva template, messy fonts)?
• Is the feed visually consistent or chaotic?
• What type of content is shown (food photos, Reels thumbnails, random lifestyle)?
• Is there a pop-up or login overlay blocking the content? (If yes, note it but still analyze what IS visible)

Evaluate this account using BOTH the screenshot AND the page text.

Return ONLY valid JSON — no markdown, no code fences, no extra text:
{{
  "qualifies": true or false,
  "disqualify_reason": "reason string if false, else null",
  "pain_points_triggered": [list of criterion numbers that apply, e.g. [1, 4]],
  "followers": approximate follower count as integer or null,
  "last_post_approx": "how long ago was the last post, e.g. '3 weeks ago' or null",
  "pain_point": "1-2 sentence personalized pain point in Simplified Chinese if qualifies=true, else null"
}}

The pain_point field MUST:
- Be Simplified Chinese, 1-2 sentences max — short, direct, specific
- Reference what you ACTUALLY SAW in the screenshot (e.g. specific visual observations)
- Include concrete data from page text where available (follower count, last post date, like counts)
- Use NO second-person address (no 您/你/你们). State facts only.
- NOT be a generic template

Good examples (style reference only — write based on actual observations):
- "Feed 以手机随拍为主，构图随意、光线不均，没有任何设计海报。"
- "最近一次发帖是三个月前，粉丝 1,200，帖子互动基本为零。"
- "有发食物海报，但设计明显是套 Canva 模板，字体混乱，整体视觉很廉价。"
- "Feed 全是 Reels，没有任何食物照片或海报，粉丝仅 800 人。"
- "每周都在更新，但图片都是随手拍，打光和构图都很粗糙，没有品牌感。"
"""

# Search fallback prompt: used when Playwright scraping fails
ICP_SEARCH_PROMPT = """
You are a social media analyst for a cold email lead generation campaign targeting F&B businesses.

QUALIFYING CRITERIA (lead qualifies if ANY apply):
1. Followers < 50,000
2. Posts roughly once a month or less
3. No new post for more than 1 month
4. Feed dominated by casual phone snapshots (unedited, random composition)
5. Feed is almost entirely Reels with NO food design posters — ONLY if followers < 20,000
6. Has food design posters but quality is poor (basic Canva templates, bad fonts/colors)
7. Rarely posts food-related content (but has posted some food content before)

DISQUALIFYING CRITERIA (exclude if ANY apply):
- Followers >= 50,000
- Feed is primarily professional photography
- Never posted any food content at all
- Feed is all Reels AND followers >= 20,000
- Not an F&B business

Use Google Search to find information about this business's social media presence.

Restaurant: {name}
Instagram: {ig_url}
Facebook: {fb_url}

Return ONLY valid JSON — no markdown, no code fences:
{{
  "qualifies": true or false,
  "disqualify_reason": "reason string if false, else null",
  "pain_points_triggered": [list of criterion numbers],
  "followers": approximate follower count as integer or null,
  "last_post_approx": "how long ago was the last post or null",
  "pain_point": "1-2 sentence personalized pain point in Simplified Chinese if qualifies=true, else null"
}}

The pain_point must be specific to this account, mention real observed details, and use no second-person address.
"""


def call_gemini_vision(prompt: str, screenshot_b64: str) -> str:
    """Send screenshot + text prompt to Gemini Vision. No search grounding."""
    payload = {
        "contents": [{
            "parts": [
                {"inline_data": {"mime_type": "image/jpeg", "data": screenshot_b64}},
                {"text": prompt},
            ]
        }],
        "generationConfig": {"temperature": 0.4, "maxOutputTokens": 16384},
    }
    r = requests.post(GEMINI_URL, json=payload, timeout=90)
    r.raise_for_status()
    data = r.json()
    candidate = data["candidates"][0]
    parts = candidate.get("content", {}).get("parts", [])
    text_parts = [p["text"] for p in parts if "text" in p]
    if not text_parts:
        raise ValueError(f"No text in vision response (finishReason={candidate.get('finishReason')})")
    return text_parts[-1].strip()


def call_gemini_search(prompt: str) -> str:
    """Call Gemini with Google Search grounding (fallback path)."""
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "tools": [{"google_search": {}}],
        "generationConfig": {"temperature": 0.4, "maxOutputTokens": 16384},
    }
    r = requests.post(GEMINI_URL, json=payload, timeout=90)
    r.raise_for_status()
    data = r.json()
    candidate = data["candidates"][0]
    parts = candidate.get("content", {}).get("parts", [])
    text_parts = [p["text"] for p in parts if "text" in p]
    if not text_parts:
        raise ValueError(f"No text in search response (finishReason={candidate.get('finishReason')})")
    return text_parts[-1].strip()


def parse_json_response(text: str) -> dict:
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    match = re.search(r"\{[\s\S]*\}", text)
    if match:
        text = match.group(0)
    return json.loads(text)


def analyze_lead(name: str, ig_url: str, fb_url: str) -> dict:
    """
    Analyze one lead.
    Primary path: Playwright scrape → Gemini Vision (screenshot + page text)
    Fallback path: Google Search grounding (if scraping fails for both URLs)
    Facebook is tried before Instagram (more accessible without login).
    """
    scrape_result = None
    url_used = None

    # Try FB first (public pages work without login), then IG
    for url in filter(None, [fb_url, ig_url]):
        scraped = scrape_page(url)
        if scraped.get("error"):
            print(f"    Scrape fail ({url[:55]}): {scraped['error']}")
            continue
        if scraped.get("screenshot_b64"):
            scrape_result = scraped
            url_used = url
            break

    # ── Path A: Vision ────────────────────────────────────────────────────
    if scrape_result and scrape_result.get("screenshot_b64"):
        platform = scrape_result["platform"]
        page_text = scrape_result.get("page_text", "")
        print(f"    Vision analysis ({platform}) — {url_used[:55]}")

        prompt = ICP_VISION_PROMPT.format(
            name=name,
            platform=platform,
            url=url_used,
            # Send first 4000 chars of page text (enough for follower/date/engagement data)
            page_text=page_text[:4000] if page_text else "(page text unavailable)",
        )

        try:
            raw = call_gemini_vision(prompt, scrape_result["screenshot_b64"])
            return parse_json_response(raw)
        except Exception as e:
            print(f"    Vision call failed: {e} — falling back to search")

    # ── Path B: Google Search fallback ────────────────────────────────────
    print(f"    Search grounding fallback")
    prompt = ICP_SEARCH_PROMPT.format(
        name=name,
        ig_url=ig_url or "N/A",
        fb_url=fb_url or "N/A",
    )

    try:
        raw = call_gemini_search(prompt)
        return parse_json_response(raw)
    except json.JSONDecodeError:
        try:
            # Last resort: search without grounding
            payload = {
                "contents": [{"parts": [{"text": prompt}]}],
                "generationConfig": {"temperature": 0.4, "maxOutputTokens": 4096},
            }
            r = requests.post(GEMINI_URL, json=payload, timeout=90)
            r.raise_for_status()
            raw2 = r.json()["candidates"][0]["content"]["parts"][0]["text"]
            return parse_json_response(raw2)
        except Exception as e2:
            return {"qualifies": None, "error": f"JSON parse failed after retry: {e2}"}
    except requests.HTTPError as e:
        detail = ""
        try:
            detail = e.response.json().get("error", {}).get("message", "")
        except Exception:
            pass
        return {"qualifies": None, "error": f"HTTP {e.response.status_code}: {detail or str(e)}"}
    except Exception as e:
        return {"qualifies": None, "error": str(e)}


# ─── Helpers ─────────────────────────────────────────────────────────────────

def safe_str(s) -> str:
    if s is None:
        return ""
    return str(s).encode("cp1252", "replace").decode("cp1252")


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) == 3:
        input_path = sys.argv[1]
        output_path = sys.argv[2]
    elif len(sys.argv) == 1:
        input_path = "leads_output/mount_austin_cafe_restaurant.xlsx"
        output_path = "leads_output/mount_austin_qualified.xlsx"
    else:
        print("Usage: py analyze_leads.py <input.xlsx> <output.xlsx>")
        sys.exit(1)

    if not os.path.exists(input_path):
        print(f"ERROR: Input file not found: {input_path}")
        sys.exit(1)

    wb_in = openpyxl.load_workbook(input_path)
    ws_in = wb_in.active
    rows = list(ws_in.iter_rows(values_only=True))
    header = rows[0]
    data_rows = [r for r in rows[1:] if r[0]]
    col_map = {v: i for i, v in enumerate(header)}

    total = len(data_rows)
    print(f"Loaded {total} leads from {input_path}")
    print(f"Columns: {list(header)}\n")

    def get_col(keys, default=None):
        for k in keys:
            if k in col_map:
                return col_map[k]
        return default

    name_col = get_col(["Restaurant Name", "name"], 0)
    ig_col   = get_col(["Instagram", "instagram", "instagram/facebook"])
    fb_col   = get_col(["Facebook", "facebook"])

    qualified = []
    errors = []

    for idx, row in enumerate(data_rows, 1):
        name   = (row[name_col] if name_col is not None else row[0]) or ""
        ig_url = (row[ig_col]   if ig_col  is not None else "") or ""
        fb_url = (row[fb_col]   if fb_col  is not None else "") or ""
        if ig_col is not None and fb_col is None:
            fb_url = ig_url

        if not ig_url and not fb_url:
            print(f"[{idx}/{total}] SKIP (no social link): {safe_str(name)}")
            continue

        print(f"[{idx}/{total}] Analyzing: {safe_str(name)}")
        result = analyze_lead(name, ig_url, fb_url)

        qualifies = result.get("qualifies")
        triggers  = result.get("pain_points_triggered", [])
        followers = result.get("followers")

        if qualifies is True:
            pain_point = result.get("pain_point", "")
            print(f"  QUALIFIED | Followers: {followers} | Triggers: {triggers}")
            print(f"  {safe_str(str(pain_point)[:120])}")
            qualified.append((*row, pain_point))
        elif qualifies is False:
            reason = result.get("disqualify_reason", "")
            print(f"  DISQUALIFIED | Followers: {followers} | {safe_str(reason)}")
        else:
            err = result.get("error", "unknown")
            print(f"  ERROR: {safe_str(str(err))}")
            errors.append((name, err))

        if idx < total:
            time.sleep(RATE_LIMIT_DELAY)

    # ── Write output Excel ────────────────────────────────────────────────────
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "Qualified Leads"

    out_header = list(header) + ["Pain Point"]
    ws_out.append(out_header)
    for cell in ws_out[1]:
        cell.font      = Font(bold=True, color="FFFFFF")
        cell.fill      = PatternFill(fill_type="solid", fgColor="2E75B6")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in qualified:
        ws_out.append(list(row))

    for col_cells in ws_out.columns:
        max_len    = max((len(str(c.value or "")) for c in col_cells), default=10)
        col_letter = openpyxl.utils.get_column_letter(col_cells[0].column)
        ws_out.column_dimensions[col_letter].width = min(max_len + 4, 80)

    pain_col_idx = len(out_header)
    for row_cells in ws_out.iter_rows(min_row=2, min_col=pain_col_idx, max_col=pain_col_idx):
        for cell in row_cells:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    os.makedirs(
        os.path.dirname(output_path) if os.path.dirname(output_path) else ".",
        exist_ok=True,
    )
    wb_out.save(output_path)

    print(f"\n{'='*50}")
    print(f"Done. {len(qualified)}/{total} leads qualified.")
    if errors:
        print(f"Errors ({len(errors)}):")
        for n, err in errors:
            print(f"  - {safe_str(n)}: {safe_str(str(err))}")
    print(f"Output: {output_path}")


if __name__ == "__main__":
    main()

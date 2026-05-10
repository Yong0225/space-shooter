# Restaurant / Cafe Lead Scraper

Scrapes Google Maps for restaurant/cafe leads in any neighborhood and outputs an Excel file with Name, Website, Email, Instagram, Facebook.

## Script

`scrap.py` — reuse this file every time. Just change the config at the top.

## How to run

```
py scrap.py              # start or resume
py scrap.py --reset      # wipe progress and restart fresh
```

## What to change for a new location / campaign

Open `scrap.py` and edit these three things at the top:

### 1. Output file name
```python
OUTPUT   = "OTR leads.xlsx"      # change to e.g. "Miami Brickell leads.xlsx"
PROGRESS = "otr_progress.json"   # change to e.g. "miami_progress.json"
```

### 2. Target count
```python
TARGET = 350    # how many leads you want
```

### 3. Search queries
Replace the `SEARCH_QUERIES` list with queries for the new area. Use 12–20 queries mixing food categories + neighborhood name to get well past TARGET unique results (Maps caps at ~60 per query).

```python
SEARCH_QUERIES = [
    "restaurants Brickell Miami Florida",
    "cafe Brickell Miami Florida",
    "coffee shop Brickell Miami Florida",
    "bar Brickell Miami Florida",
    "brewery Brickell Miami Florida",
    "brunch Brickell Miami Florida",
    "pizza Brickell Miami Florida",
    "food Brickell Miami Florida",
    "bakery dessert Brickell Miami Florida",
    "sushi Brickell Miami Florida",
    "mexican restaurant Brickell Miami Florida",
    "italian restaurant Brickell Miami Florida",
    "burger Brickell Miami Florida",
    "breakfast Brickell Miami Florida",
    "wine bar Brickell Miami Florida",
    "cocktail bar Brickell Miami Florida",
    "seafood Brickell Miami Florida",
    "vegan Brickell Miami Florida",
    "food Miami Florida 33131",       # zip code helps catch stragglers
    "restaurant Miami Florida 33131",
]
```

## Output Excel columns

| Column | What it contains |
|--------|-----------------|
| Name | Business name from Google Maps |
| Website | Business website URL |
| Email | Best email found (green row = has email) |
| Instagram | Instagram page URL |
| Facebook | Facebook page URL |

Green rows = email found. White/blue alternating = no email.

## How email is found

1. Visit the business website homepage
2. Also visit any `/contact` or `/about` sub-pages (up to 2 extra pages)
3. Scan for `mailto:` links and plain email patterns
4. If no email found but Facebook link exists → visit Facebook About page and scan there too

## Resume / crash recovery

Progress is saved to `otr_progress.json` after every single lead. If the script crashes or you close it, just run `py scrap.py` again — it picks up exactly where it stopped.

## Dependencies

```
py -m pip install playwright openpyxl
py -m playwright install chromium
```

No API keys needed. Browser runs in headed (visible) mode so you can solve CAPTCHAs manually if Google blocks.

## Typical timings

| Phase | Time |
|-------|------|
| Phase 1 – collect places from Maps | ~10–15 min (12–20 queries) |
| Phase 2 – scrape each website | ~30–45 min for 350 leads |
| Total | ~45–60 min for 350 leads |

## Filtering: alcohol-only venues

Places are automatically skipped if their name contains any of these phrases (case-insensitive):

`cocktail bar`, `cocktail lounge`, `cocktail room`, `wine bar`, `wine lounge`, `whiskey bar`, `whiskey lounge`, `bourbon bar`, `spirits bar`, `spirits lounge`, `speakeasy`

Restaurants and bars that also serve food (e.g. "bar & grill", "sports bar", "brewery") are **not** filtered — only venues that are clearly alcohol/cocktail focused.

To add more skip terms, edit `ALCOHOL_SKIP_KEYWORDS` in `scrap.py`.

## Tips

- If Maps returns fewer places than expected (< 50 per query), Google may be rate-limiting. Wait 10–15 min and re-run — progress is saved.
- Add zip codes as extra queries (e.g. `"food Cincinnati Ohio 45202"`) to catch places that don't show up under the neighborhood name.
- The `--reset` flag deletes both the progress JSON and the Excel file and starts completely fresh.

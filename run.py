import argparse
import sys
from urllib.parse import urlparse
from core.runner import run
from core.rules import BROWSERS

ap = argparse.ArgumentParser(description="Excel test runner (Playwright)")
ap.add_argument("excel", help="Path of the Excel file, e.g. input/test_cases.xlsx")
ap.add_argument("--url", required=True, help="Website URL, e.g. https://www.saucedemo.com")
ap.add_argument("--browser", default="chromium", choices=BROWSERS)
ap.add_argument("--headed", action="store_true", help="Show the browser window")
ap.add_argument("--timeout", type=float, default=10, help="Wait time in seconds (default 10)")
ap.add_argument("--no-video", action="store_true", help="Do not record video")
ap.add_argument("--no-screenshot", action="store_true", help="Do not take screenshot on fail")
ap.add_argument("--username", default="", help="Value for {{USERNAME}} in the Excel steps")
ap.add_argument("--password", default="", help="Value for {{PASSWORD}} in the Excel steps")
ap.add_argument("--only", default="", help="Run only these test cases, e.g. TC-001,TC-003")
a = ap.parse_args()

u = urlparse(a.url)
if u.scheme not in ("http", "https") or not u.netloc:
    ap.error("--url must be a valid URL starting with http:// or https://")

sys.exit(run(a.excel, {
    "url": a.url.rstrip("/"),
    "browser": a.browser,
    "headless": not a.headed,
    "timeout": a.timeout,
    "video": not a.no_video,
    "screenshot": not a.no_screenshot,
    "username": a.username,
    "password": a.password,
    "only": [x.strip() for x in a.only.split(",") if x.strip()],
}))

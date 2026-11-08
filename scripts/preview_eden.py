"""
Take a clean preview screenshot of /eden — login as admin first so user sees
the actual dashboard, not the login page.
"""
import os
import time
from playwright.sync_api import sync_playwright

OUT_DIR = "/home/z/my-project/download/eden_screenshots"
os.makedirs(OUT_DIR, exist_ok=True)

BASE_URL = "http://127.0.0.1:8000"
EMAIL = "admin@example.com"
PASSWORD = "testing"

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    context = browser.new_context(
        viewport={"width": 1440, "height": 900},
        device_scale_factor=1.5,
        locale="en-US",
    )
    page = context.new_page()

    # Login
    print("→ Login as admin@example.com")
    page.goto(f"{BASE_URL}/eden/default/user/login", wait_until="networkidle", timeout=30000)
    page.wait_for_timeout(500)
    page.fill('input[name="email"]', EMAIL)
    page.fill('input[name="password"]', PASSWORD)
    page.click('input[type="submit"]')
    page.wait_for_load_state("networkidle", timeout=30000)
    page.wait_for_timeout(1500)
    print(f"  After login → URL: {page.url}, title: {page.title()}")

    # Screenshot: Eden home (viewport only)
    out_home = f"{OUT_DIR}/preview_eden_home.png"
    page.screenshot(path=out_home, full_page=False)
    print(f"  ✓ {out_home}  ({os.path.getsize(out_home)//1024} KB)")

    # Also screenshot Organizations page (a more interesting view)
    page.goto(f"{BASE_URL}/eden/org/organisation", wait_until="networkidle", timeout=30000)
    page.wait_for_timeout(1000)
    out_org = f"{OUT_DIR}/preview_eden_organisations.png"
    page.screenshot(path=out_org, full_page=False)
    print(f"  ✓ {out_org}  ({os.path.getsize(out_org)//1024} KB)")

    # And the Map page (visually striking)
    page.goto(f"{BASE_URL}/eden/gis", wait_until="networkidle", timeout=30000)
    page.wait_for_timeout(2000)  # let map tiles load
    out_map = f"{OUT_DIR}/preview_eden_map.png"
    page.screenshot(path=out_map, full_page=False)
    print(f"  ✓ {out_map}  ({os.path.getsize(out_map)//1024} KB)")

    browser.close()

print()
print("Done.")

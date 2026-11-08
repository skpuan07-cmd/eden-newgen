"""
Browse Sahana Eden via headless Chromium:
- Login as admin
- Screenshot home page + key module pages
- Save PNGs to /home/z/my-project/download/eden_screenshots/
"""
import os
import time
from playwright.sync_api import sync_playwright

OUT_DIR = "/home/z/my-project/download/eden_screenshots"
os.makedirs(OUT_DIR, exist_ok=True)

BASE_URL = "http://127.0.0.1:8000"
EMAIL = "admin@example.com"
PASSWORD = "testing"

pages_to_shoot = [
    ("01_home",            "/eden/default/index",     "Eden Home (Dashboard)"),
    ("02_org_organisation","/eden/org/organisation",  "Organization Registry"),
    ("03_org_office",      "/eden/org/office",         "Offices"),
    ("04_pr_person",       "/eden/pr/person",          "Person Registry"),
    ("05_hrm_staff",       "/eden/hrm/staff",          "Staff (HRM)"),
    ("06_inv_warehouse",   "/eden/inv/warehouse",      "Warehouses (Inventory)"),
    ("07_project_project", "/eden/project/project",    "Projects"),
    ("08_asset_asset",     "/eden/asset/asset",        "Assets"),
    ("09_req_req",         "/eden/req/req",            "Requests"),
    ("10_cms_post",        "/eden/cms/post",           "CMS Posts"),
    ("11_gis_index",       "/eden/gis",                "GIS (Map)"),
    ("12_default_config",  "/eden/admin/config",       "Eden Configuration"),
]

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    context = browser.new_context(
        viewport={"width": 1440, "height": 900},
        device_scale_factor=1.5,
        locale="en-US",
    )
    page = context.new_page()

    # Capture console errors for debugging
    console_errors = []
    page.on("console", lambda msg: console_errors.append(f"[{msg.type}] {msg.text}") if msg.type in ("error", "warning") else None)

    print("=" * 60)
    print("Sahana Eden — Browsing Session")
    print("=" * 60)
    print()
    print(f"Viewport: 1440x900 (1.5x device scale)")
    print(f"User:     {EMAIL}")
    print()

    # Step 1: Go to login page
    print("Step 1: Open login page")
    page.goto(f"{BASE_URL}/eden/default/user/login", wait_until="networkidle", timeout=30000)
    page.wait_for_timeout(800)
    print(f"  Title: {page.title()}")
    print(f"  URL:   {page.url}")

    # Step 2: Fill login form
    print()
    print("Step 2: Fill login form")
    page.fill('input[name="email"]', EMAIL)
    page.fill('input[name="password"]', PASSWORD)
    page.wait_for_timeout(300)

    # Step 3: Submit login
    print()
    print("Step 3: Submit login")
    page.click('input[type="submit"]')
    page.wait_for_load_state("networkidle", timeout=30000)
    page.wait_for_timeout(1500)
    print(f"  After login:")
    print(f"  Title: {page.title()}")
    print(f"  URL:   {page.url}")

    # Step 4: Take screenshots of all pages
    print()
    print("Step 4: Take screenshots of key pages")
    print(f"  {'Page':<35} {'HTTP':<6} {'Title':<30} {'Screenshot'}")
    print(f"  {'-'*35} {'-'*6} {'-'*30} {'-'*30}")

    for fname, path, desc in pages_to_shoot:
        try:
            response = page.goto(f"{BASE_URL}{path}", wait_until="networkidle", timeout=30000)
            page.wait_for_timeout(1000)  # let JS settle
            status = response.status if response else "?"
            title = page.title()[:30]
            out = f"{OUT_DIR}/{fname}.png"
            page.screenshot(path=out, full_page=False)
            print(f"  {path:<35} {status:<6} {title:<30} ✓ {fname}.png")
        except Exception as e:
            print(f"  {path:<35} ERROR  {str(e)[:50]:<30}")

    # Step 5: One full-page screenshot of the home page (scrolling)
    print()
    print("Step 5: Full-page screenshot of Eden home (scrolling)")
    try:
        page.goto(f"{BASE_URL}/eden/default/index", wait_until="networkidle", timeout=30000)
        page.wait_for_timeout(1500)
        out = f"{OUT_DIR}/00_home_fullpage.png"
        page.screenshot(path=out, full_page=True)
        print(f"  ✓ {out}")
    except Exception as e:
        print(f"  ✗ {e}")

    # Step 6: Save menu structure (extract from home page)
    print()
    print("Step 6: Extract top menu structure from Eden home")
    try:
        page.goto(f"{BASE_URL}/eden/default/index", wait_until="networkidle", timeout=30000)
        page.wait_for_timeout(1000)
        # Get all menu items
        menu_items = page.eval_on_selector_all(
            'nav a, .top-bar a, #menu a, ul.menu a, [role="navigation"] a',
            '''els => els.map(e => ({text: e.textContent.trim(), href: e.getAttribute('href')}))
                     .filter(m => m.text && m.href && m.href.startsWith('/eden/'))
                     .slice(0, 40)'''
        )
        # Deduplicate by href
        seen = set()
        unique = []
        for m in menu_items:
            if m["href"] not in seen:
                seen.add(m["href"])
                unique.append(m)
        print(f"  Found {len(unique)} unique menu items:")
        for m in unique[:30]:
            print(f"    {m['text'][:30]:<32}  →  {m['href']}")
        # Save to file
        import json
        with open(f"{OUT_DIR}/menu_structure.json", "w") as f:
            json.dump(unique, f, indent=2, ensure_ascii=False)
        print(f"  ✓ Saved to {OUT_DIR}/menu_structure.json")
    except Exception as e:
        print(f"  ✗ {e}")

    browser.close()

print()
print(f"All screenshots saved to: {OUT_DIR}")
print()
print(f"Console errors/warnings captured: {len(console_errors)}")
for e in console_errors[:10]:
    print(f"  {e[:100]}")

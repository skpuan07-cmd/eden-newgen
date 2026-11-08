#!/usr/bin/env python3
"""Test Sahana Eden installation: login + visit key endpoints."""
import re
import sys
import time
import urllib.parse
import urllib.request
import urllib.error
import http.cookiejar
import json

BASE_URL = "http://127.0.0.1:8000"
EMAIL = "admin@example.com"
PASSWORD = "testing"

def make_request(url, method="GET", data=None, cookies=None, referer=None, allow_redirects=True):
    """Make an HTTP request, return (status, headers, body)."""
    if isinstance(data, dict):
        data = urllib.parse.urlencode(data).encode("utf-8")
    elif isinstance(data, str):
        data = data.encode("utf-8")

    req = urllib.request.Request(url, data=data, method=method)
    if cookies is not None:
        cookies.add_cookie_header(req)
    if referer:
        req.add_header("Referer", referer)
    req.add_header("User-Agent", "EdenInstallTest/1.0")

    # Disable redirects for inspection
    class NoRedirectHandler(urllib.request.HTTPRedirectHandler):
        def redirect_request(self, req, fp, code, msg, headers, newurl):
            if not allow_redirects:
                return None
            return super().redirect_request(req, fp, code, msg, headers, newurl)

    opener = urllib.request.build_opener(NoRedirectHandler)
    try:
        resp = opener.open(req, timeout=30)
        body = resp.read()
        if cookies is not None:
            cookies.extract_cookies(resp, req)
        return resp.status, dict(resp.headers), body
    except urllib.error.HTTPError as e:
        body = e.read()
        if cookies is not None:
            cookies.extract_cookies(e, req)
        return e.code, dict(e.headers), body

def extract_form(html):
    """Extract _formname and _formkey from a web2py form."""
    fm = re.search(rb'name="_formname"\s+[^>]*value="([^"]+)"', html)
    fk = re.search(rb'name="_formkey"\s+[^>]*value="([^"]+)"', html)
    formname = fm.group(1).decode() if fm else ""
    formkey = fk.group(1).decode() if fk else ""
    return formname, formkey

def extract_title(html):
    m = re.search(rb"<title>([^<]*)</title>", html)
    return m.group(1).decode("utf-8", errors="replace") if m else "(no title)"

def main():
    cookies = http.cookiejar.CookieJar()

    print("=" * 60)
    print("Sahana Eden — Installation Test")
    print("=" * 60)
    print()

    # 1) Wait for server
    print("Step 1: Wait for server to be ready...")
    for i in range(30):
        try:
            status, _, _ = make_request(f"{BASE_URL}/eden", cookies=cookies)
            if status < 500:
                print(f"  ✓ Server ready (after {i}s, status={status})")
                break
        except Exception:
            time.sleep(1)
    else:
        print("  ✗ Server did not come up after 30s")
        sys.exit(1)
    print()

    # 2) GET login form
    print("Step 2: GET /eden/default/user/login")
    status, _, body = make_request(f"{BASE_URL}/eden/default/user/login", cookies=cookies)
    print(f"  HTTP {status}, size={len(body)}b, title='{extract_title(body)}'")
    formname, formkey = extract_form(body)
    print(f"  _formname='{formname}', _formkey={formkey[:20]}...")
    print()

    # 3) POST login
    print(f"Step 3: POST login ({EMAIL} / {PASSWORD})")
    login_data = {
        "email": EMAIL,
        "password": PASSWORD,
        "_formname": formname,
        "_formkey": formkey,
    }
    status, headers, body = make_request(
        f"{BASE_URL}/eden/default/user/login",
        method="POST",
        data=login_data,
        cookies=cookies,
        referer=f"{BASE_URL}/eden/default/user/login",
        allow_redirects=False,
    )
    print(f"  HTTP {status}, size={len(body)}b")
    if status in (303, 302):
        print(f"  Location: {headers.get('Location', '(none)')}")
    elif status == 200:
        # Check if it's the login form (failed) or a logged-in page
        title = extract_title(body)
        if "login" in title.lower():
            print(f"  ⚠ Login failed — title still '{title}'")
        else:
            print(f"  Login appears successful (title='{title}')")
    print()

    # 4) Test endpoints (with cookies — should now be logged in)
    test_endpoints = [
        ("/eden/default/index",            "Eden home"),
        ("/eden",                          "Eden root"),
        ("/eden/org/organisation",         "Organization Registry"),
        ("/eden/org/office",               "Offices"),
        ("/eden/org/facility",             "Facilities"),
        ("/eden/pr/person",                "Person Registry"),
        ("/eden/hrm/staff",                "Staff (HRM)"),
        ("/eden/hrm/volunteer",            "Volunteers (HRM)"),
        ("/eden/inv/warehouse",            "Warehouses (Inventory)"),
        ("/eden/inv/inv_item",             "Inventory Items"),
        ("/eden/project/project",          "Projects"),
        ("/eden/asset/asset",              "Assets"),
        ("/eden/req/req",                  "Requests"),
        ("/eden/cap/alert",                "CAP Alerts"),
        ("/eden/dvr/dvregistry",           "Disaster Victim Registry"),
        ("/eden/cms/post",                 "CMS Posts"),
        ("/eden/default/index",            "Eden home (revisit)"),
    ]

    print("Step 4: Test endpoints after login")
    print(f"  {'Endpoint':<35} {'HTTP':<6} {'Size':<10} {'Title'}")
    print(f"  {'-'*35} {'-'*6} {'-'*10} {'-'*40}")
    results = []
    for path, name in test_endpoints:
        try:
            status, _, body = make_request(f"{BASE_URL}{path}", cookies=cookies, allow_redirects=False)
            title = extract_title(body) if body else "(empty)"
            size = f"{len(body)}b"
            print(f"  {path:<35} {status:<6} {size:<10} {title[:50]}")
            results.append((path, status, len(body), title))
        except Exception as e:
            print(f"  {path:<35} ERROR: {e}")
            results.append((path, "ERR", 0, str(e)))
    print()

    # 5) JSON API
    print("Step 5: Test JSON API endpoints (with auth)")
    json_endpoints = [
        "/eden/org/organisation.json",
        "/eden/pr/person.json",
        "/eden/project/project.json",
        "/eden/inv/warehouse.json",
    ]
    print(f"  {'Endpoint':<45} {'HTTP':<6} {'Size':<10} First 80 chars")
    print(f"  {'-'*45} {'-'*6} {'-'*10} {'-'*40}")
    for path in json_endpoints:
        try:
            status, _, body = make_request(f"{BASE_URL}{path}", cookies=cookies)
            preview = body[:80].decode("utf-8", errors="replace").replace("\n", " ")
            print(f"  {path:<45} {status:<6} {len(body)}b    {preview}")
        except Exception as e:
            print(f"  {path:<45} ERROR: {e}")
    print()

    # 6) Database tables check
    print("Step 6: Database table count")
    import sqlite3
    db_path = "/home/z/my-project/eden_install/web2py/applications/eden/databases/storage.sqlite"
    try:
        conn = sqlite3.connect(db_path)
        tables = conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"
        ).fetchall()
        print(f"  ✓ {len(tables)} tables in storage.sqlite")
        # Sample some interesting tables
        interesting = ["auth_user", "auth_group", "org_organisation", "org_organisation_type",
                       "pr_person", "hrm_human_resource", "inv_warehouse", "inv_inv_item",
                       "project_project", "asset_asset", "cap_alert", "s3_permission"]
        print(f"  Sample tables:")
        for t in interesting:
            count = conn.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
            print(f"    {t:<30}  {count} rows")
        # Count users
        users = conn.execute("SELECT email, first_name FROM auth_user").fetchall()
        print(f"  Demo users in auth_user:")
        for u in users:
            print(f"    {u[1]:<20} <{u[0]}>")
        conn.close()
    except Exception as e:
        print(f"  ✗ Database check failed: {e}")
    print()

    # Summary
    print("=" * 60)
    print("Summary")
    print("=" * 60)
    failed = [r for r in results if r[1] != 200]
    if not failed:
        print(f"✓ ALL {len(results)} endpoints responded HTTP 200")
    else:
        print(f"⚠ {len(failed)}/{len(results)} endpoints did NOT return HTTP 200:")
        for path, status, size, title in failed:
            print(f"    {path}: HTTP {status} ({title[:40]})")
    print()

if __name__ == "__main__":
    main()

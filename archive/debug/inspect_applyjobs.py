"""Inspect the Apply Jobs page after login."""
from playwright.sync_api import sync_playwright
import time, json

URL = "https://firststepjob.com/application-incharge/apply-jobs?userId=google-oauth2|110590764416739071400"

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False, slow_mo=150)
    page = browser.new_context(viewport={"width": 1280, "height": 900}).new_page()

    print("[*] Loading page...")
    page.goto(URL, wait_until="networkidle", timeout=30000)
    time.sleep(3)

    print(f"[*] Current URL: {page.url}")
    print(f"[*] Page title: {page.title()}")

    # Dump visible buttons and links
    print("\n====== BUTTONS ======")
    btns = page.evaluate("""() => {
        const out = [];
        document.querySelectorAll('button, a').forEach((el, i) => {
            const txt = el.innerText?.trim().slice(0, 60);
            const href = el.href || '';
            if (txt) out.push({i, tag: el.tagName, text: txt, href: href.slice(0, 80)});
        });
        return out;
    }""")
    for b in btns:
        print(json.dumps(b))

    # Dump page HTML snippet
    print("\n====== PAGE HTML (first 5000 chars) ======")
    print(page.content()[:5000])

    print("\n[*] Done. Closing in 5s...")
    time.sleep(5)
    browser.close()

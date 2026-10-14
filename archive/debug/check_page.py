from playwright.sync_api import sync_playwright
import time

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    context = browser.contexts[0]
    page = context.pages[0] if context.pages else context.new_page()
    print(f"URL: {page.url}")
    print(f"Title: {page.title()}")
    html = page.content()
    print(f"HTML snippet: {html[:2000]}")
    browser.close()

from playwright.sync_api import sync_playwright
with sync_playwright() as p:
    b = p.chromium.connect_over_cdp("http://127.0.0.1:9222")
    ctx = b.contexts[0]
    for pg in ctx.pages:
        u = pg.url
        if "chatgpt" in u:
            print("CHATGPT url:", u[:60])
            box = pg.locator('#prompt-textarea, div[contenteditable="true"], textarea').count()
            print("  prompt box present:", box)
        if "spreadsheets" in u:
            print("SHEET url:", u[:60])
            nb = pg.locator('input.waffle-name-box, [aria-label="Name box"]').count()
            print("  name box present:", nb)

"""Re-extract the response already on the ChatGPT tab with the FIXED logic and
report the bold words. (No new prompt — reuse what's on screen.)"""
import re
import importlib
import run_all
importlib.reload(run_all)
from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    b = p.chromium.connect_over_cdp("http://127.0.0.1:9222")
    ctx = b.contexts[0]
    page = next((pg for pg in ctx.pages if "chatgpt.com" in pg.url), None)
    if not page:
        print("No ChatGPT tab."); raise SystemExit
    ans = run_all._extract_chatgpt_answer(page) or ""
    print("===== EXTRACTED (fixed logic) =====")
    print(ans[:1000])
    print("\n** marker count:", ans.count("**"))
    print("any **** (bad):", "****" in ans)
    bolds = re.findall(r'\*\*(.+?)\*\*', ans)
    print("bold words captured:", bolds)

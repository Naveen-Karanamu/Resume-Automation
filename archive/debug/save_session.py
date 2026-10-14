"""
Run this ONCE to save your Google login session.

Steps:
  1. Run this script — it will open a normal Chrome window
  2. Log in with Google manually in that window
  3. Once you're on the Apply Jobs page, come back here and press Enter
  4. Session is saved to session.json

Run:
    python3.10 save_session.py
"""

import json
import shutil
import subprocess
import time
from pathlib import Path
from playwright.sync_api import sync_playwright

URL          = "https://firststepjob.com/application-incharge/apply-jobs?userId=google-oauth2|110590764416739071400"
SESSION_FILE = "session.json"
CHROME_EXE   = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
DEBUG_PORT   = 9222
TEMP_DIR     = r"C:\Users\karan\AppData\Local\Temp\ChromeDebugSession"

# ── Step 1: Launch a normal Chrome with remote debugging (no automation flags) ──
print("[*] Launching Chrome with remote debugging...")
Path(TEMP_DIR).mkdir(parents=True, exist_ok=True)

chrome_proc = subprocess.Popen([
    CHROME_EXE,
    f"--remote-debugging-port={DEBUG_PORT}",
    f"--user-data-dir={TEMP_DIR}",
    "--no-first-run",
    "--no-default-browser-check",
    URL,
])

print(f"[*] Chrome launched (PID {chrome_proc.pid})")
print("\n[!] A Chrome window opened. Log in with Google manually.")
print("[!] Once you are on the Apply Jobs page, come back here and press Enter.")
input("\nPress Enter when ready to save session...")

# ── Step 2: Connect Playwright to the running Chrome and extract cookies ──
with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp(f"http://localhost:{DEBUG_PORT}")
    context = browser.contexts[0]

    storage = context.storage_state()
    with open(SESSION_FILE, "w") as f:
        json.dump(storage, f)

    print(f"\n[+] Session saved to {SESSION_FILE}")
    browser.close()

# ── Step 3: Close Chrome ──────────────────────────────────────────────────────
chrome_proc.terminate()
print("[+] Done. You can now run apply_jobs.py without logging in again.")

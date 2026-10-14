"""
ChatGPT -> extract Summary / Resp 1 / Resp 2 -> append to resume_data.xlsx
=========================================================================
Steps:
  1. Open ChatGPT (reusing the running Chrome / logged-in session)
  2. Paste the prompt from chatgpt_prompt.txt and send it
  3. Wait for the full response
  4. Extract:
       - Professional Summary        -> Summary   (column C)
       - Responsibilities 1          -> Ex Desc 1 (column D)
       - Responsibilities 2          -> Ex Desc 2 (column E)
  5. Append a new row to resume_data.xlsx

Reads : chatgpt_prompt.txt
Writes: resume_data.xlsx (appends one row), chatgpt_response.txt (raw reply)

Run:
    python chatgpt_to_excel.py
"""

import os
import re
import subprocess
import time
import urllib.request
from pathlib import Path

import openpyxl
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# ── CONFIG ──────────────────────────────────────────────────────────────────
CHATGPT_URL  = "https://chatgpt.com/"
CHROME_EXE   = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
DEBUG_PORT   = 9222
HOME         = Path(os.path.expanduser("~"))
TEMP_DIR     = str(HOME / "AppData" / "Local" / "Temp" / "ChromeDebugSession")
PROMPT_IN    = Path(__file__).parent / "chatgpt_prompt.txt"
RESP_OUT     = Path(__file__).parent / "chatgpt_response.txt"
EXCEL_FILE   = Path(__file__).parent / "resume_data.xlsx"
# ─────────────────────────────────────────────────────────────────────────────


def debug_port_up():
    try:
        urllib.request.urlopen(f"http://127.0.0.1:{DEBUG_PORT}/json", timeout=2)
        return True
    except Exception:
        return False


def launch_chrome():
    print("[*] Launching Chrome with remote debugging...")
    Path(TEMP_DIR).mkdir(parents=True, exist_ok=True)
    proc = subprocess.Popen([
        CHROME_EXE,
        f"--remote-debugging-port={DEBUG_PORT}",
        f"--user-data-dir={TEMP_DIR}",
        "--no-first-run", "--no-default-browser-check", "about:blank",
    ])
    for _ in range(30):
        if debug_port_up():
            print(f"[+] Chrome ready (PID {proc.pid}).")
            return proc
        time.sleep(1)
    raise RuntimeError("Chrome debug port did not come up.")


def clean(text):
    """Strip markdown bold/formatting markers but keep bullet content."""
    if not text:
        return ""
    text = re.sub(r'\*\*(.*?)\*\*', r'\1', text)   # **bold** -> bold
    text = re.sub(r'^#{1,6}\s*', '', text, flags=re.MULTILINE)
    return text.strip()


def extract_sections(text):
    """
    Pull the three sections out of the ChatGPT reply. Headings may appear as
    'Professional Summary', '**Professional Summary**', 'Summary:', etc., and
    the sections run until the next heading.
    """
    # Normalise bold markers around headings so regex can find them.
    norm = re.sub(r'\*\*', '', text)

    def grab(start_patterns, stop_patterns):
        for sp in start_patterns:
            m = re.search(sp, norm, re.IGNORECASE)
            if not m:
                continue
            start = m.end()
            # find the earliest stop heading after start
            stop_idx = len(norm)
            for xp in stop_patterns:
                sm = re.search(xp, norm[start:], re.IGNORECASE)
                if sm:
                    stop_idx = min(stop_idx, start + sm.start())
            return norm[start:stop_idx].strip(" :\n\r-")
        return ""

    summary = grab(
        [r'Professional Summary\s*[:\-]?', r'\bSummary\s*[:\-]?'],
        [r'Responsibilit(?:y|ies)\s*1', r'Responsibilit(?:y|ies)\b'],
    )
    resp1 = grab(
        [r'Responsibilit(?:y|ies)\s*1\s*[:\-]?'],
        [r'Responsibilit(?:y|ies)\s*2'],
    )
    resp2 = grab(
        [r'Responsibilit(?:y|ies)\s*2\s*[:\-]?'],
        [r'\Z'],
    )
    return clean(summary), clean(resp1), clean(resp2)


def main():
    print("=" * 64)
    print("  ChatGPT -> Summary / Resp1 / Resp2 -> append to resume_data.xlsx")
    print("=" * 64)

    if not PROMPT_IN.exists():
        print("[!] chatgpt_prompt.txt not found. Run the sheets step first.")
        return
    prompt = PROMPT_IN.read_text(encoding="utf-8").strip()
    # The sheet wraps the whole prompt in quotes — strip a single enclosing pair.
    if prompt.startswith('"') and prompt.endswith('"'):
        prompt = prompt[1:-1].strip()
    print(f"[+] Loaded prompt ({len(prompt)} chars).")

    chrome_proc = None if debug_port_up() else launch_chrome()
    if debug_port_up() and chrome_proc is None:
        print("[+] Reusing the already-running Chrome debug session.")

    with sync_playwright() as p:
        print("[*] Connecting to Chrome over CDP...")
        browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{DEBUG_PORT}")
        context = browser.contexts[0]

        # Reuse an existing ChatGPT tab if one is open, else open a new one.
        page = None
        for pg in context.pages:
            if "chatgpt.com" in pg.url or "chat.openai.com" in pg.url:
                page = pg
                print("[+] Reusing an open ChatGPT tab.")
                break
        if page is None:
            page = context.new_page()
            print("[*] Opening ChatGPT...")
            page.goto(CHATGPT_URL, wait_until="domcontentloaded", timeout=60000)
        page.bring_to_front()
        try:
            page.wait_for_load_state("networkidle", timeout=15000)
        except PlaywrightTimeoutError:
            pass
        time.sleep(3)

        if "auth" in page.url.lower() or "login" in page.url.lower():
            print("[!] ChatGPT needs a login. Sign in in the Chrome window, then re-run.")
            browser.close()
            if chrome_proc:
                chrome_proc.terminate()
            return

        # ── Step 2: paste the prompt and send ────────────────────────────
        try:
            box = page.locator('#prompt-textarea, div[contenteditable="true"], textarea').first
            box.wait_for(state="visible", timeout=20000)
            box.click()
            # ProseMirror contenteditable: fill then dispatch, fallback to type.
            try:
                box.fill(prompt)
            except Exception:
                page.keyboard.insert_text(prompt)
            time.sleep(0.5)
            print("[+] Prompt entered. Sending...")
            page.keyboard.press("Enter")
        except Exception as e:
            print(f"[!] Could not enter the prompt: {str(e)[:120]}")
            browser.close()
            if chrome_proc:
                chrome_proc.terminate()
            return

        # ── Step 3: wait for the response to finish streaming ────────────
        print("[*] Waiting for ChatGPT to respond...")
        time.sleep(5)
        done = False
        for _ in range(150):  # up to ~2.5 min
            # The send/stop button toggles; when streaming stops, the stop
            # button disappears. Also check that an assistant message exists.
            streaming = page.locator(
                'button[data-testid="stop-button"], button[aria-label*="Stop" i]'
            ).count() > 0
            has_msg = page.locator('[data-message-author-role="assistant"]').count() > 0
            if has_msg and not streaming:
                done = True
                break
            time.sleep(1)
        time.sleep(2)
        if not done:
            print("[~] Response may not have fully finished; capturing what's there.")

        # ── Step 4: read the last assistant message ──────────────────────
        msgs = page.locator('[data-message-author-role="assistant"]').all()
        if not msgs:
            print("[!] No assistant response found.")
            browser.close()
            if chrome_proc:
                chrome_proc.terminate()
            return
        reply = msgs[-1].inner_text()
        RESP_OUT.write_text(reply, encoding="utf-8")
        print(f"[+] Response captured ({len(reply)} chars). Saved to {RESP_OUT.name}.")

        summary, resp1, resp2 = extract_sections(reply)
        print(f"\n  Summary   : {len(summary)} chars")
        print(f"  Resp 1    : {len(resp1)} chars")
        print(f"  Resp 2    : {len(resp2)} chars")

        if not (summary or resp1 or resp2):
            print("[!] Could not parse any section. Check chatgpt_response.txt and "
                  "adjust the headings. Not appending to Excel.")
            browser.close()
            if chrome_proc:
                chrome_proc.terminate()
            return

        # ── Step 5: append a row to resume_data.xlsx ─────────────────────
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
        # Columns: A=Job Title, B=Job Location, C=Summary, D=Ex Desc 1,
        #          E=Ex Desc 2, F=Resume Name
        ws.append(["", "", summary, resp1, resp2, ""])
        wb.save(EXCEL_FILE)
        print(f"\n[+] Appended a new row to {EXCEL_FILE.name} "
              f"(now {ws.max_row} rows).")

        print("\n" + "─" * 64)
        print("PROFESSIONAL SUMMARY")
        print("─" * 64)
        print(summary[:600])
        print("\n" + "─" * 64)
        print("RESPONSIBILITIES 1 (Ex Desc 1)")
        print("─" * 64)
        print(resp1[:600])
        print("\n" + "─" * 64)
        print("RESPONSIBILITIES 2 (Ex Desc 2)")
        print("─" * 64)
        print(resp2[:600])
        print("─" * 64)

        time.sleep(2)
        browser.close()
        if chrome_proc:
            chrome_proc.terminate()
        print("\n[+] Done.")


if __name__ == "__main__":
    main()

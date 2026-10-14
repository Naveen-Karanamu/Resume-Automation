"""Run resume_automation copy.py's full flow for ONE row (row 1), using its own
fresh browser exactly like the real script, and report the result + PDF path."""
import importlib.util, time, re
from playwright.sync_api import sync_playwright

spec = importlib.util.spec_from_file_location("rac", "resume_automation copy.py")
rac = importlib.util.module_from_spec(spec); spec.loader.exec_module(rac)

email, password, max_pages = rac.read_credentials()
rows = rac.read_excel()
row = rows[0]
print(f"[*] Row 1: job_title='{row['job_title']}' resume='{row['resume_name']}'")
print(f"[*] ex_desc_1 has ** markers: {'**' in row['ex_desc_1']}")
print(f"[*] max_pages from credentials: {max_pages}")

with sync_playwright() as p:
    browser = p.chromium.launch(headless=rac.HEADLESS, slow_mo=rac.SLOW_MO)
    ctx = browser.new_context(accept_downloads=True,
                              viewport={"width": 1280, "height": 900})
    page = ctx.new_page()
    try:
        rac.login(page, email, password)
        rac.load_last_resume(page)
        rac.update_resume(page, row, max_pages=max_pages)

        # verify real bold in the builder before downloading
        descs = page.locator('div.mt-1.px-3.py-2.block.w-full.rounded-md.border.border-gray-300[contenteditable]').all()
        for idx, label in ((1, "Ex Desc 1"), (3, "Ex Desc 2")):
            if idx < len(descs):
                html = descs[idx].evaluate("e => e.innerHTML")
                text = descs[idx].evaluate("e => e.innerText")
                bolds = [a or c for a, c in re.findall(r'<b>(.*?)</b>|<strong>(.*?)</strong>', html)]
                print(f"\n{label}: literal '**' present = {'**' in text}")
                print(f"  bold runs ({len(bolds)}):", bolds[:10])

        rac.download_resume(page, row["resume_name"] or row["job_title"] or "test_row1")
    except Exception as e:
        print("[ERROR]", str(e)[:150])
    finally:
        time.sleep(2)
        ctx.close()
        browser.close()

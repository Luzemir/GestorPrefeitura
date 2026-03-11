from playwright.sync_api import sync_playwright
import time

def test_export():
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp('http://localhost:9222')
        page = browser.contexts[0].pages[0]
        try:
            print("Filling start date...")
            page.fill('input[id$="idStart_input"]', "01/02/2026")
            print("Filling end date...")
            page.fill('input[id$="idEnd_input"]', "28/02/2026")
            
            print("Clicking Pesquisar...")
            pesquisar = page.locator('a:has-text("Pesquisar")').first
            pesquisar.click()
            time.sleep(4)
            
            print("Clicking XLS...")
            with page.expect_download(timeout=60000) as download_info:
                xls_btn = page.locator('a:has-text("XLS")').first
                if not xls_btn.is_visible():
                    print("XLS not visible directly. Taking screenshot of page.")
                    page.screenshot(path="test_error.png")
                xls_btn.click()
                
            download = download_info.value
            filepath = download.suggested_filename
            print(f"Download successful! Suggested filename: {filepath}")
        except Exception as e:
            print(f"Error: {e}")

if __name__ == "__main__":
    test_export()

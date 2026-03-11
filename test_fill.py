from playwright.sync_api import sync_playwright
import time

def test_livro():
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp('http://localhost:9222')
        page = browser.contexts[0].pages[0]
        try:
            print("URL: ", page.url)
            print("Filling start date...")
            page.fill('input[id$="idStart_input"]', "01/02/2026")
            print("Filling end date...")
            page.fill('input[id$="idEnd_input"]', "28/02/2026")
            
            print("Finding Excel button...")
            btn = page.locator('button:has-text("Excel")').first
            if btn.is_visible():
                print("Excel button found!")
            else:
                print("Excel button not visible. Trying span with text Excel instead...")
                btn2 = page.locator('span:has-text("Excel")').first
                if btn2.is_visible():
                    print("Span Excel found.")
                else:
                    print("Could not find any visible text Excel.")
                    
            print("Trying to click Gerenciar using evaluate fallback...")
            res = page.evaluate('''() => {
                let items = Array.from(document.querySelectorAll('span.nav-label'));
                let gerenciar = items.find(e => e.innerText && e.innerText.includes('Gerenciar'));
                if(gerenciar) { gerenciar.click(); return 'Clicked via JS'; }
                return 'Not Found';
            }''')
            print("Evaluate result:", res)
                
            print("Done test.")
        except Exception as e:
            print(f"Error: {e}")

if __name__ == "__main__":
    test_livro()

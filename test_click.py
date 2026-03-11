from playwright.sync_api import sync_playwright
import time

def test_clicks():
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp('http://localhost:9222')
        page = browser.contexts[0].pages[0]
        print(f'Active Page: {page.url}')
        
        try:
            gerenciar = page.locator('span.nav-label:has-text("Gerenciar NFSe")').first
            print(f'Gerenciar NFSe visible? {gerenciar.is_visible()}')
            gerenciar.click(timeout=5000)
            print('Clicked Gerenciar NFSe')
            time.sleep(1)
            
            livro = page.locator('span.nav-label:has-text("Livro Fiscal")').first
            print(f'Livro Fiscal visible? {livro.is_visible()}')
            livro.click(timeout=5000)
            print('Clicked Livro Fiscal')
            time.sleep(2)
            
            print('Done. Active URL:', page.url)
        except Exception as e:
            print('Error clicking:', e)

if __name__ == '__main__':
    test_clicks()

from playwright.sync_api import sync_playwright

def get_page_content():
    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp("http://localhost:9222")
            context = browser.contexts[0]
            
            # Print page titles to identify the active one
            for idx, page in enumerate(context.pages):
                title = page.title()
                print(f"[{idx}] {title} - {page.url}")
                if "bemVindo" in page.url or "notafiscal/paginas" in page.url:
                    content = page.content()
                    with open(f"page_dump_{idx}.html", "w", encoding="utf-8") as f:
                        f.write(content)
                    print(f"Saved page {idx} HTML to page_dump_{idx}.html")
        except Exception as e:
            print(f"Error connecting: {e}")

if __name__ == "__main__":
    get_page_content()

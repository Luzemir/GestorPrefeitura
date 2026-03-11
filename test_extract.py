from playwright.sync_api import sync_playwright

def find_export():
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp('http://localhost:9222')
        page = browser.contexts[0].pages[0]
        try:
            out = page.evaluate('''() => {
                let items = Array.from(document.querySelectorAll('a, button, input[type=submit], input[type=button]'));
                return items.map(e => {
                    let html = (e.innerHTML || "").replace(/\\n/g, ' ').replace(/\\s+/g, ' ');
                    return e.tagName + " | id: " + e.id + " | class: " + e.className + " | html: " + html;
                });
            }''')
            with open('buttons_dump.txt', 'w', encoding='utf-8') as f:
                for x in out: f.write(x + '\\n')
            print("Dumped buttons to buttons_dump.txt. Total:", len(out))
        except Exception as e:
            print("Error:", e)

if __name__ == "__main__":
    find_export()

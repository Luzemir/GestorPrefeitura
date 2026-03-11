from playwright.sync_api import sync_playwright

def take_shot():
    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp("http://localhost:9222")
            context = browser.contexts[0]
            if not context.pages:
                print("No pages open!")
                return
            page = context.pages[0]
            print(f"Active Page: {page.url}")
            page.screenshot(path="tela_debug.png")
            print("Saved tela_debug.png")
            
            # Print specifically the active frames to know if there's a frame mismatch
            for idx, frame in enumerate(page.frames):
                print(f"Frame {idx}: {frame.name} - {frame.url}")
        except Exception as e:
            print(f"Error: {e}")

if __name__ == "__main__":
    take_shot()

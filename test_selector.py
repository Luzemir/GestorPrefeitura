from playwright.sync_api import sync_playwright

print('Iniciando Teste Rapido...')
with sync_playwright() as p:
    try:
        browser = p.chromium.connect_over_cdp('http://localhost:9222')
        page = browser.contexts[0].pages[0]
        print('URL atual:', page.url)
        frames = page.frames
        print('Total Frames:', len(frames))
        for f in frames:
            print('  Frame URL:', f.url)
            
        print('Tenta achar input cnpj na main page (timeout 5s)...')
        loc = page.locator('input[id$="idInputMask"]').first
        
        try:
            loc.wait_for(timeout=5000, state="visible")
            print('Achei o input e esta visivel!')
        except Exception as e:
            print('Nao está visivel... Verificando iframes')
            
            for f in frames:
                loc_frame = f.locator('input[id$="idInputMask"]').first
                try:
                    loc_frame.wait_for(timeout=2000, state="visible")
                    print(f'ACHEI NO FRAME: {f.name} / {f.url}')
                except:
                    pass
    except Exception as e:
        print('Erro no script de debug:', e)

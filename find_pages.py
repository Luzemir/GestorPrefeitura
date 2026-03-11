from playwright.sync_api import sync_playwright

print('Iniciando Teste de Multipaginas...')
with sync_playwright() as p:
    try:
        browser = p.chromium.connect_over_cdp('http://localhost:9222')
        context = browser.contexts[0]
        
        print(f'Total de abas abertas no Chrome: {len(context.pages)}')
        for index, pg in enumerate(context.pages):
            print(f'Aba {index} - URL: {pg.url}')
            
            # Testa cada aba
            loc = pg.locator('input[id*="idCpfCnpj"]').first
            try:
                if loc.is_visible(timeout=3000):
                    print('>>> ACHEI O INPUT NESTA ABA!!!!! <<<')
            except Exception as e:
                print('...input nao encontrado nesta aba')
            
    except Exception as e:
        print('Erro:', e)

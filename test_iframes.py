import time
from playwright.sync_api import sync_playwright

print('Iniciando Teste de JS Injection...')
with sync_playwright() as p:
    try:
        browser = p.chromium.connect_over_cdp('http://localhost:9222')
        page = browser.contexts[0].pages[0]
        
        print('Injetando um JS pra descobrir iframes e campos Cnpj:')
        js_cmd = """
        () => {
            const inputs = Array.from(document.querySelectorAll('input'));
            const iframes = Array.from(document.querySelectorAll('iframe'));
            return {
                num_inputs: inputs.length,
                inputs_ids: inputs.map(i => i.id).filter(id => id.includes('cnpj') || id.includes('Cnpj')),
                num_iframes: iframes.length,
                iframes_names: iframes.map(f => f.name || f.id)
            };
        }
        """
        resultado = page.evaluate(js_cmd)
        print("Resultado do DOM da Main Page:", resultado)
        
        # Testar dentro do iframe se houver (Prefeitura de CG costuma usar iframes)
        if resultado['num_iframes'] > 0:
            for i in range(len(page.frames)):
                if i == 0: continue # main frame
                frame = page.frames[i]
                try:
                    res_frame = frame.evaluate(js_cmd)
                    print(f"Resultado dentro do frame {frame.name}:", res_frame)
                except Exception as e:
                    print(f"Nao foi possivel acessar o dom do frame {i}: {e}")
            
    except Exception as e:
        print('Erro geral:', e)

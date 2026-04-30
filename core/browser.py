from contextlib import contextmanager
from playwright.sync_api import sync_playwright, Page
import sys

@contextmanager
def get_browser_page():
    """
    Context manager que conecta a uma instância local do Google Chrome e retorna a primeira aba aberta.
    O Chrome deve estar rodando com a flag --remote-debugging-port=9222
    """
    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp('http://localhost:9222')
            context = browser.contexts[0]
            if not context.pages:
                raise Exception("Nenhuma aba aberta no Chrome.")
            yield context.pages[0]
        except Exception as e:
            print(f"Erro ao conectar com o Chrome: {e}")
            print("Certifique-se de iniciar o Chrome com --remote-debugging-port=9222")
            raise

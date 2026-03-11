import os
import time
from playwright.sync_api import sync_playwright

# Setup caminhos
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LIVROS_DIR = os.path.join(BASE_DIR, "livros")
USER_DATA_DIR = os.path.join(BASE_DIR, "chrome_profile")

if not os.path.exists(LIVROS_DIR):
    os.makedirs(LIVROS_DIR)

# URL do portal
URL_LOGIN = "https://nfse.campogrande.ms.gov.br/notafiscal/paginas/portal/index.html#/login"

def run_login_flow():
    print("Iniciando Playwright...")
    with sync_playwright() as p:
        # Iniciamos com navegador CHROMIUM (ou msedge se o certificado estiver lá)
        # headless=False para que vejamos a tela
        # Usamos launch_persistent_context para manter os cookies da sessao caso o certificado pessa
        browser = p.chromium.launch_persistent_context(
            user_data_dir=USER_DATA_DIR,
            headless=False,
            channel="chrome", # Força usar o Chrome do SO, útil para certificados instalados
            args=["--start-maximized"]
        )
        
        page = browser.pages[0] if browser.pages else browser.new_page()
        print(f"Acessando url: {URL_LOGIN}")
        page.goto(URL_LOGIN)
        
        print("\n" + "="*50)
        print("MOMENTO DE AÇÃO MANUAL NECESSÁRIA:")
        print("1. O navegador foi aberto.")
        print("2. Faça o login com seu Certificado Digital.")
        print("3. QUANDO VOCÊ ESTIVER NA TELA DE 'PESQUISAR CNPJ', volte no terminal e aperte ENTER.")
        print("="*50 + "\n")
        
        input("Pressione ENTER após concluir o login manual na página de selecionar empresa...")
        
        print("Continuando a automação...")
        # Aqui ficará o código da Fase da Empresa
        
        browser.close()

if __name__ == "__main__":
    run_login_flow()

import os
import time
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from playwright.sync_api import sync_playwright

# Setup caminhos
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(BASE_DIR)
LIVROS_DIR = os.path.join(PROJECT_DIR, "livros")
USER_DATA_DIR = os.path.join(PROJECT_DIR, "chrome_profile")
CONFIG_FILE = os.path.join(PROJECT_DIR, "planejamento", "Relação de empresas.xlsx")
MASTER_FILE = os.path.join(PROJECT_DIR, "consolidado.xlsx")

# Gera o nome do arquivo de log unico para esta execucao
TIMESTAMP_EXECUCAO = datetime.now().strftime("%Y%m%d-%H%M%S")
LOG_FILE = os.path.join(LIVROS_DIR, f"log_execucao_{TIMESTAMP_EXECUCAO}.txt")

def log_exec(cnpj, nome, comp, status, detalhes):
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    linha = f"[{agora}] CNPJ: {cnpj} | Nome: {nome} | Comp: {comp} | Status: {status} | Info: {detalhes}\n"
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(linha)

for d in [LIVROS_DIR, USER_DATA_DIR]:
    if not os.path.exists(d):
        os.makedirs(d)

URL_LOGIN = "https://nfse.campogrande.ms.gov.br/notafiscal/paginas/portal/index.html#/login"

def load_master_df():
    if os.path.exists(MASTER_FILE):
        return pd.read_excel(MASTER_FILE)
    return pd.DataFrame(columns=[
        "Mes_Competencia", "CNPJ", "Nome", "Qtd_Notas_Ativas", "Qtd_Notas_Canceladas", 
        "Faturamento_Bruto", "Valor_ISS", "Valor_INSS", "Valor_IR", "Valor_COFINS", "Valor_CSLL", "Valor_PIS", "Valor_Liquido"
    ])

def save_master_df(df):
    df.to_excel(MASTER_FILE, index=False)

def read_targets():
    if not os.path.exists(CONFIG_FILE):
        print(f"Arquivo {CONFIG_FILE} nao existe!")
        return []
    # Lendo o excel base da prefeitura - pulando cabeçalho (linha 1 parece ser vazia/titulo)
    df = pd.read_excel(CONFIG_FILE, header=1)
    
    targets = []
    # As colunas parecem ser "Razão Social" e "CNPJ" ou algo similar.
    # Vamos iterar pelas linhas. Coluna 1 = Nome, Coluna 2 = CNPJ
    for _, row in df.iterrows():
        nome = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        cnpj_raw = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
        
        # Limpar mascara do CNPJ (Manter só numeros)
        cnpj_clean = "".join(filter(str.isdigit, cnpj_raw))
        
        if len(cnpj_clean) == 14: # Apenas cnpjs validos
            targets.append({
                'Nome_Empresa': nome,
                'CNPJ': cnpj_clean
            })
            
    return targets

def process_livro_fiscal(filepath, target, competencia):
    print(f"Processando planilha: {filepath}")
    
    if filepath.lower().endswith('.pdf'):
        print(f"Erro: O arquivo baixado é um PDF ({filepath}). O menu Saída falhou em selecionar Excel. Pulando processamento Pandas...")
        return None
        
    # O excel baixa com algumas linhas de cabecalho
    try:
        df = pd.read_excel(filepath, header=5)
    except Exception as e:
        print(f"Erro ao ler {filepath}: {e}")
        return None
        
    df_ativas = df[df['SITUACAO'].str.lower() == 'ativa']
    df_Canceladas = df[df['SITUACAO'].str.lower() == 'cancelada']
    
    qtd_ativas = len(df_ativas)
    qtd_canceladas = len(df_Canceladas)
    
    fat_bruto = df_ativas['VALOR NF'].sum() if 'VALOR NF' in df_ativas.columns else 0
    iss = df_ativas['VALOR ISS'].sum() if 'VALOR ISS' in df_ativas.columns else 0
    inss = df_ativas['VALOR INSS'].sum() if 'VALOR INSS' in df_ativas.columns else 0
    ir = df_ativas['VALOR IR'].sum() if 'VALOR IR' in df_ativas.columns else 0
    cof = df_ativas['VALOR COFINS'].sum() if 'VALOR COFINS' in df_ativas.columns else 0
    csl = df_ativas['VALOR CSLL'].sum() if 'VALOR CSLL' in df_ativas.columns else 0
    pis = df_ativas['VALOR PIS'].sum() if 'VALOR PIS' in df_ativas.columns else 0
    liq = df_ativas['VALOR LÍQUIDO'].sum() if 'VALOR LÍQUIDO' in df_ativas.columns else 0
    
    resumo = {
        "Mes_Competencia": competencia,
        "CNPJ": target['CNPJ'],
        "Nome": target.get('Nome_Empresa', ''),
        "Qtd_Notas_Ativas": qtd_ativas,
        "Qtd_Notas_Canceladas": qtd_canceladas,
        "Faturamento_Bruto": fat_bruto,
        "Valor_ISS": iss,
        "Valor_INSS": inss,  
        "Valor_IR": ir,
        "Valor_COFINS": cof,
        "Valor_CSLL": csl,
        "Valor_PIS": pis,
        "Valor_Liquido": liq
    }
    return resumo

def run_bot(competencia_mes_ano):
    targets = read_targets()
    if not targets:
        return
        
    master_df = load_master_df()
    
    # Verifica ja processados neste mes
    already_processed = set()
    if not master_df.empty:
        mask = master_df["Mes_Competencia"] == competencia_mes_ano
        already_processed = set(master_df[mask]["CNPJ"].astype(str).tolist())

    print("Iniciando Playwright (Conectando ao Chrome aberto)...")
    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp('http://localhost:9222')
            default_context = browser.contexts[0]
            page = default_context.pages[0] if default_context.pages else default_context.new_page()
            
            print(f"Conectado ao Chrome! Navegando para a página inicial...")
            page.goto(URL_LOGIN)
            
        except Exception as e:
            print(f"ERRO: Não foi possível conectar ao Chrome. Certifique-se de abri-lo pelo atalho de depuração.\nDetalhes: {e}")
            return
            
        input("Pressione ENTER SOMENTE quando você estiver na tela de 'Pesquisar CNPJ' após o Certificado Digital...")
        
        for target in targets:
            cnpj = target['CNPJ'].strip()
            
            if cnpj in already_processed:
                print(f"CNPJ {cnpj} já processado na competência {competencia_mes_ano}. Pulando...")
                continue
                
            print(f">>> Iniciando target: {cnpj}")
            nome_empresa = target.get('Nome_Empresa', '')
            
            
            # Funcao auxiliar para clicar usando javascript puro ignorando opacidade/visibilidade do JSF
            def js_click_text(selector, text):
                page.evaluate(f'''() => {{
                    Array.from(document.querySelectorAll("{selector}")).forEach(el => {{
                        if((el.innerText || el.textContent || "").includes("{text}") || (el.title || "").includes("{text}")) {{
                            el.click();
                        }}
                    }});
                }}''')
            
            try:
                # Voltar para home / pesquisa clicando no Menu "Seleciona Cadastro"
                # Abre o sidebar se estiver fechado (o botao Menu tem uma div com texto Menu e class hide-menu ou logo-menu)
                # Abre o sidebar se estiver fechado (o botao Menu tem uma div com texto Menu e class hide-menu ou logo-menu)
                try:
                    menu_btn = page.locator('.header-link.hide-menu').first
                    if menu_btn.is_visible():
                        menu_btn.click(timeout=1000)
                        time.sleep(1)
                except:
                    pass
                
                # Clica em 'Seleciona Cadastro' no menu lateral
                # Clica em 'Seleciona Cadastro' no menu lateral
                js_click_text('span.nav-label', 'Seleciona Cadastro')
                time.sleep(2)
            except Exception as e:
                print(f"Erro ao clicar em Seleciona Cadastro no menu: {e}")
                
            try:
                # 1. Informar o CNPJ e pesquisar
                page.fill('input[id*="idCpfCnpj"]', cnpj)
                js_click_text('a', 'Pesquisar')
                time.sleep(3)
                
                # Clicar na linha da tabela onde ta o CNPJ/check (Como o CNPJ vem formatado com mascara na web, clicamos no primeiro result)
                has_selecionar = page.evaluate('''() => {
                    let links = Array.from(document.querySelectorAll("a"));
                    return links.some(el => (el.textContent || "").includes("Selecionar") || (el.title || "").includes("Selecionar"));
                }''')
                
                if not has_selecionar:
                    print(f"CNPJ {cnpj} não listado nas procurações. Pulando pros próximos...")
                    log_exec(cnpj, nome_empresa, competencia_mes_ano, "IGNORADO", "CNPJ não encontrado na tabela de procurações após pesquisa")
                    continue

                js_click_text('a', 'Selecionar')
                time.sleep(4)
                
                try: # Se houver accordion ou submenu lateral "Gerenciar NFSE" -> "Livro Fiscal"
                    js_click_text('span.nav-label', 'Gerenciar NFSe')
                    time.sleep(1)
                        
                    # Clica especificamente em Livro Fiscal no menu
                    js_click_text('span.nav-label', 'Livro Fiscal')
                    time.sleep(2)
                except Exception as e:
                    print(f"Nota: Erro ao clicar no menu Livro Fiscal: {e}")
                    page.screenshot(path="error_menu.png")
                    time.sleep(2)
                
                # 3. Informar Datas e Filtros no form
                # O site so precisa que digitemos MM/AAAA (e.g. 022026), e a mascara processa o dia
                data_raw = competencia_mes_ano.replace("/", "")
                
                inp_ini = page.locator('input[id$="idStart_input"]').first
                inp_ini.clear()
                inp_ini.type(data_raw, delay=50)
                
                inp_end = page.locator('input[id$="idEnd_input"]').first
                inp_end.clear()
                inp_end.type(data_raw, delay=50)
                
                # Marcar Todos situacao e Exigibilidade (Super Primefaces Checkbox Hack)
                page.evaluate('''() => {
                    // Clica na caixa estilizada visual do Primefaces caso não esteja checada
                    document.querySelectorAll('.ui-chkbox-box:not(.ui-state-active)').forEach(el => el.click());
                    
                    // Força o checkbox escondido internamente para 'true' por segurança
                    document.querySelectorAll('input[type="checkbox"]').forEach(el => {
                        if (!el.checked) {
                            el.checked = true;
                            el.dispatchEvent(new Event('change', { bubbles: true }));
                        }
                    });
                }''')
                        
                # Clica Pesquisar para carregar a tabela!
                js_click_text('a', 'Pesquisar')
                time.sleep(4)
                        
                # Tentar encontrar a Saída "Excel" antes de clicar no export / gerar
                # Tentar encontrar a Saída "Excel" antes de clicar no export / gerar
                try:
                    # 1. Clicar na caixa de dropdown Saída via Playwright para garantir que o evento Focus/Click ocorra no DOM real
                    lbl_saida = page.locator('label', has_text='Saída').first
                    if lbl_saida.is_visible():
                        # Clica na seta apontadora (trigger) do select correspondente
                        lbl_saida.locator('xpath=..').locator('.ui-selectonemenu-trigger').click(force=True)
                        time.sleep(1)
                        
                        # 2. Clicar na opção Excel na lista renderizada no fim do HTML
                        item_excel = page.locator('li.ui-selectonemenu-item:has-text("Excel"), li.ui-selectonemenu-item:has-text("XLS")').first
                        if item_excel.is_visible():
                            item_excel.click(force=True)
                        else:
                            js_click_text('li.ui-selectonemenu-item', 'Excel')
                        time.sleep(1)
                except Exception as e:
                    print(f"Nota: não foi possível mudar a 'Saída' para Excel pelo dropdown. Continuando... Erro: {e}")

                # 4. Baixar Excel
                print("Aguardando geracao do excel ou mensagem de erro (vazio)...")
                
                # Encontra o botão de XLS ou Exportar/Gerar
                btn_excel = None
                for selector in [
                    'a:has-text("Download")',
                    'button:has-text("Download")',
                    'button:has-text("Excel")', 
                    'a:has-text("XLS")', 
                    'button:has-text("Gerar")',
                    'button:has-text("Imprimir")',
                    'a.ui-commandlink:has(i.pi-file-excel)'
                ]:
                    element = page.locator(selector).first
                    if element.is_visible():
                        btn_excel = element
                        break
                        
                if not btn_excel:
                    print("Botão de Excel/XLS não localizado visivelmente. Capturando imagem da tela e forçando click no último recurso...")
                    page.screenshot(path="error_excel_btn.png")
                    btn_excel = page.locator('button[type="submit"]').first
                    
                # Configura a captura do download em background
                download_obj = [None]
                def handle_download(d):
                    download_obj[0] = d
                    
                page.on("download", handle_download)
                btn_excel.click()
                
                # Loop de espera (Polling) para verificar se o download iniciou OU se deu erro de "Nenhum registro"
                has_error = False
                timeout_count = 0
                while timeout_count < 120:
                    if download_obj[0]:
                        break
                        
                    # Verifica a mensagem de erro vermelha no painel informando que nao tem notas
                    if page.get_by_text("Nenhum registro", exact=False).first.is_visible():
                        has_error = True
                        break
                            
                    page.wait_for_timeout(1000) # Aguarda 1 segundo processando a fila de eventos do Playwright
                    timeout_count += 1
                    
                # Limpa o listener para não acumular multiplos manipuladores ao longo das 300+ empresas
                page.remove_listener("download", handle_download)
                
                if has_error:
                    print(f"Empresa {cnpj} não possui notas na competência {competencia_mes_ano}. Preenchendo zerado e pulando para a próxima...")
                    resumo_vazio = {
                        "Mes_Competencia": competencia_mes_ano,
                        "CNPJ": target['CNPJ'],
                        "Nome": target.get('Nome_Empresa', ''),
                        "Qtd_Notas_Ativas": 0,
                        "Qtd_Notas_Canceladas": 0,
                        "Faturamento_Bruto": 0,
                        "Valor_ISS": 0,
                        "Valor_INSS": 0,  
                        "Valor_IR": 0,
                        "Valor_COFINS": 0,
                        "Valor_CSLL": 0,
                        "Valor_PIS": 0,
                        "Valor_Liquido": 0
                    }
                    master_df = pd.concat([master_df, pd.DataFrame([resumo_vazio])], ignore_index=True)
                    save_master_df(master_df)
                    log_exec(cnpj, nome_empresa, competencia_mes_ano, "SUCESSO_VAZIO", "Sem notas fiscais no período. Valores zerados registrados.")
                    continue # Pula o processamento do excel e vai pra rotina do proximo cnpj
                    
                if not download_obj[0]:
                    raise Exception("Timeout aguardando download ou mensagem de erro da prefeitura.")
                    
                download = download_obj[0]
                
                # Resiliencia de nome gerado
                sugg_ext = download.suggested_filename.split('.')[-1]
                if not sugg_ext: sugg_ext = 'xlsx'
                
                # Previne que barras ou caracteres ilegais no nome da empresa (ex: S/S) criem pastas indesejadas
                nome_seguro = nome_empresa.replace('/', '-').replace('\\', '-').strip()
                if not nome_seguro:
                    nome_seguro = 'empresa'
                    
                filepath = os.path.join(LIVROS_DIR, f"{competencia_mes_ano.replace('/','')}_{nome_seguro}.{sugg_ext}")
                if os.path.exists(filepath):
                    agora = datetime.now().strftime("%H%M%S")
                    filepath = os.path.join(LIVROS_DIR, f"{competencia_mes_ano.replace('/','')}_{nome_seguro}_{agora}.{sugg_ext}")
                    
                download.save_as(filepath)
                print(f"Download salvo em: {filepath}")
                
                # 5. Processamento Local (Pandas)
                resumo = process_livro_fiscal(filepath, target, competencia_mes_ano)
                if resumo:
                    master_df = pd.concat([master_df, pd.DataFrame([resumo])], ignore_index=True)
                    save_master_df(master_df)
                    print(f"Planilha mestre atualizada com {cnpj}.")
                    log_exec(cnpj, nome_empresa, competencia_mes_ano, "SUCESSO_ARQUIVO", f"Arquivo {sugg_ext.upper()} consolidado")
                    
            except Exception as e:
                print(f"Erro ao processar Empresa {cnpj}: {e}")
                log_exec(cnpj, nome_empresa, competencia_mes_ano, "ERRO", str(e).replace('\n', ' '))
                page.screenshot(path=f"error_empresa_{cnpj}.png")
                
        print("Finalizado processamento de todas as empresas da lista!")
        browser.close()

if __name__ == "__main__":
    competencia = input("Digite a competência para buscar os livros (Formato MM/AAAA): ")
    run_bot(competencia)

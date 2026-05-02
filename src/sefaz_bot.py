import time
import os
import pandas as pd
from playwright.sync_api import sync_playwright

def run_sefaz_bot(comp, filepath, output_dir=None, log_callback=None):
    if log_callback:
        log_callback(f"Iniciando integração com a Sefaz para a competência {comp}")
        log_callback(f"Lendo empresas do arquivo: {filepath}")

    # --- Task 1.2: Leitura do arquivo base de empresas ---
    try:
        df_empresas = pd.read_excel(filepath)
        colunas_necessarias = ["Código Domínio", "Nome Empresa", "CNPJ", "Inscrição Estadual"]
        for col in colunas_necessarias:
            if col not in df_empresas.columns:
                raise Exception(f"Coluna obrigatória não encontrada: {col}")
        
        # Limpando possíveis espaços em branco nos nomes das colunas
        df_empresas.columns = df_empresas.columns.str.strip()
    except Exception as e:
        raise Exception(f"Erro na leitura da planilha de empresas: {str(e)}")

    if log_callback:
        log_callback("Conectando ao navegador...")

    # --- Fase 2: Automação Core - Autenticação e Segurança ---
    with sync_playwright() as p:
        try:
            # Conecta na porta 9223, definida na GUI para a Sefaz
            browser = p.chromium.connect_over_cdp('http://localhost:9223')
            context = browser.contexts[0]
            
            if not context.pages:
                page = context.new_page()
            else:
                page = context.pages[0]
            
            # Task 2.1: Navegar e aguardar o login manual
            if "sefaz.ms.gov.br" not in page.url:
                page.goto("https://eservicos.sefaz.ms.gov.br/")
                
            if log_callback:
                log_callback("Aguardando login manual com Certificado Digital (Até 60s)...")
            
            # Task 2.3: Validação de Segurança - Espera o nome aparecer na tela
            try:
                # Usa regex ignorando maiúsculas/minúsculas e pega o primeiro para evitar erro de Strict Mode
                elemento_nome = page.locator("text=/LUZEMIR MARTINS BARBOSA/i").first
                elemento_nome.wait_for(state="visible", timeout=60000)
                if log_callback:
                    log_callback("Acesso autorizado: LUZEMIR MARTINS BARBOSA identificado.")
            except Exception as e:
                raise Exception("Perfil 'LUZEMIR MARTINS BARBOSA' não encontrado na tela. Verifique o certificado digital e tente novamente.")

            # Task 2.2: Tratar o fechamento de possíveis pop-ups (ex: botão OK)
            try:
                if log_callback:
                    log_callback("Verificando se há mensagens/popups iniciais...")
                # Busca pelo botão OK que aparece na modal e aguarda até 5 segundos
                btn_ok = page.get_by_role("button", name="OK").first
                btn_ok.wait_for(state="visible", timeout=5000)
                btn_ok.click()
                if log_callback:
                    log_callback("Popup inicial identificado e fechado com sucesso (OK clicado).")
                time.sleep(1) # Aguarda a animação de saída do popup
            except Exception:
                if log_callback:
                    log_callback("Nenhum popup bloqueante encontrado, seguindo o fluxo.")
                pass # Se não houver popup, segue normalmente
            
            if log_callback:
                log_callback("Fase 2 concluída com sucesso! Iniciando varredura das empresas (Fase 3).")
            
            # --- Fase 3: Navegação do Bot - Mapeamento e Iteração ---
            total_empresas = len(df_empresas)
            for index, row in df_empresas.iterrows():
                nome_empresa = row.get("Nome Empresa", f"Empresa_{index}")
                ie_empresa = str(row.get("Inscrição Estadual", "")).strip()
                
                # Pula se a IE for nula
                if not ie_empresa or ie_empresa == "nan":
                    if log_callback:
                        log_callback(f"⚠️ {nome_empresa}: Inscrição Estadual inválida. Pulando...")
                    continue
                    
                if log_callback:
                    log_callback("==================================================")
                    log_callback(f"🏢 Processando ({index+1}/{total_empresas}): {nome_empresa}")
                    log_callback(f"IE: {ie_empresa}")
                    
                try:
                    # Garantir que estamos na tela inicial limpa antes de processar nova empresa
                    if index > 0:
                        page.goto("https://eservicos.sefaz.ms.gov.br/Home/PaginaInicial")
                        time.sleep(2)
                        
                    # Se a Sefaz renderizar o pop-up de aviso novamente ao recarregar/focar
                    try:
                        btn_ok = page.get_by_role("button", name="OK").first
                        if btn_ok.is_visible(timeout=2000):
                            btn_ok.click()
                            time.sleep(1)
                    except Exception:
                        pass
                        
                    # Tratamento para o Modal de Cookies da OneTrust
                    try:
                        btn_cookies = page.locator("text=/Confirmar as minhas escolhas/i").first
                        if btn_cookies.is_visible(timeout=2000):
                            btn_cookies.click()
                            time.sleep(1)
                    except Exception:
                        pass

                    # Task 3.2: Selecionar Perfil Procurador e Inscrição Estadual
                    if log_callback:
                        log_callback("Configurando o Tipo de Perfil e Contribuinte...")
                    
                    # 1. Selecionar o Tipo de Perfil "Procurador"
                    # Na imagem, o 'Tipo do Perfil' é a primeira caixa (combobox) e o 'Contribuinte' é a segunda.
                    try:
                        # Tenta encontrar e clicar na primeira caixa de seleção (Tipo de Perfil)
                        combo_perfil = page.get_by_role("combobox").nth(0)
                        if combo_perfil.is_visible():
                            combo_perfil.click()
                        else:
                            # Tenta encontrar pela label se o HTML for bem estruturado
                            page.get_by_label("Tipo do Perfil", exact=False).first.click()
                            
                        time.sleep(1)
                        # Digita 'Procurador' no campo focado e dá Enter
                        page.keyboard.type("Procurador", delay=100)
                        time.sleep(1)
                        page.keyboard.press("Enter")
                    except Exception:
                        pass # Pode já vir selecionado
                        
                    time.sleep(1)
                    
                    # 2. Inserir Inscrição Estadual em 'Contribuinte'
                    try:
                        # Tenta encontrar e clicar na segunda caixa de seleção (Contribuinte)
                        combo_contrib = page.get_by_role("combobox").nth(1)
                        if combo_contrib.is_visible():
                            combo_contrib.click()
                        else:
                            # Ou tenta clicar no placeholder que vimos na imagem
                            page.get_by_text("Selecione um contribuinte", exact=False).first.click()
                            
                        time.sleep(1)
                        # Digita a IE pausadamente para forçar o AJAX a buscar
                        page.keyboard.type(ie_empresa, delay=100)
                        time.sleep(2) # Aguarda a lista de sugestões abrir
                        
                        # Clica na opção que apareceu correspondente à IE digitada
                        # Usa regex pegando o último (o que aparece na caixa suspensa por cima)
                        opcao_dropdown = page.locator(f"text=/{ie_empresa}/i").last
                        opcao_dropdown.wait_for(state="visible", timeout=10000)
                        opcao_dropdown.click()
                        
                        if log_callback:
                            log_callback(f"✅ Empresa selecionada no portal com sucesso!")
                    except Exception as e:
                        if log_callback:
                            log_callback(f"❌ Falha ao encontrar e preencher Contribuinte: {e}")
                        raise e
                    
                    time.sleep(1)
                    
                    # --- Fase 4: Extração NFe (Emitidas e Recebidas) ---
                    if log_callback:
                        log_callback("Acessando o serviço de Informações Fiscais...")
                        
                    # Aguarda os cards de serviços aparecerem
                    time.sleep(2)
                    
                    # 1. Clicar no card "Informações Fiscais" (o botão verde na lista de serviços)
                    card_info = page.locator("text=/Informações Fiscais/i").first
                    card_info.wait_for(state="visible", timeout=15000)
                    
                    if log_callback:
                        log_callback("Clicando em Informações Fiscais (Aguardando pop-up/módulo)...")
                    
                    abas_antes = len(context.pages)
                    
                    # Tenta capturar a abertura de um pop-up (nova janela real do navegador)
                    try:
                        with context.expect_page(timeout=8000) as nova_aba_info:
                            card_info.click()
                        page_modulo = nova_aba_info.value
                        page_modulo.wait_for_load_state("domcontentloaded")
                        if log_callback: log_callback("Módulo detectado em nova janela (Pop-up).")
                    except Exception:
                        # Se não abriu imediatamente, espera mais um pouco
                        time.sleep(3)
                        if len(context.pages) > abas_antes:
                            page_modulo = context.pages[-1]
                            if log_callback: log_callback("Módulo aberto em nova janela (carregamento lento).")
                        else:
                            # Se não abriu nova aba, assume que é um Modal (sobreposição na mesma página)
                            page_modulo = page
                            if log_callback: log_callback("Módulo detectado na mesma janela (Modal HTML).")
                            
                            # Verifica se o Chrome pode ter bloqueado o pop-up
                            if log_callback: log_callback("AVISO: Se nada acontecer, verifique se o Chrome bloqueou o Pop-up na barra de endereços!")
                    
                    page_modulo.bring_to_front()
                    time.sleep(2) # Pequeno atraso para garantir animações de modal/popup
                    
                    # Task 4.1: Acesso ao submenu (Navegação pelo Menu Superior)
                    try:
                        if log_callback: log_callback("Acessando Menu de Opções > Doc Fiscais > NF-e...")
                        
                        # 1. Clica no "Menu de Opções"
                        menu_opcoes = page_modulo.locator("text=/Menu de Opções/i").first
                        menu_opcoes.wait_for(state="visible", timeout=10000)
                        menu_opcoes.click()
                        time.sleep(1.5) # Aguarda a animação do menu abrir
                        
                        # 2. Hover/Click em "Documentos Fiscais Eletrônicos"
                        # Usando regex para cobrir variações de acento (Eletronicos vs Eletrônicos)
                        doc_fiscais = page_modulo.locator("text=/Documentos Fiscais Eletr/i").first
                        doc_fiscais.hover() # Em muitos menus drop-down antigos, o hover é necessário
                        time.sleep(1)
                        doc_fiscais.click()
                        time.sleep(1)
                        
                        # 3. Clica em "Nota Fiscal Eletrônica"
                        nfe_menu = page_modulo.locator("text=/Nota Fiscal Eletr/i").first
                        nfe_menu.click()
                            
                    except Exception as e:
                        if log_callback: log_callback(f"Erro fatal ao navegar pelo Menu de Opções: {e}")
                        raise e
                    
                    if log_callback: log_callback("Acessando tela de pesquisa de notas...")
                    
                    # Aguarda o redirecionamento / carregamento da página de pesquisa
                    page_modulo.wait_for_load_state("domcontentloaded")
                    time.sleep(3)
                    
                    mes_comp, ano_comp = comp.split('/')
                    mes = int(mes_comp)
                    ano = int(ano_comp)
                    
                    import calendar
                    ultimo_dia = calendar.monthrange(ano, mes)[1]
                    data_inicial = f"01/{mes:02d}/{ano}"
                    data_final = f"{ultimo_dia:02d}/{mes:02d}/{ano}"
                    
                    # Listas para armazenar os dados capturados nesta empresa
                    nfe_emitidas = []
                    nfe_recebidas = []
                    
                    # Função auxiliar interna para capturar tabela e tratar paginação
                    def extrair_tabela_paginada(tipo_nota):
                        dados = []
                        if log_callback:
                            log_callback(f"Iniciando pesquisa de Notas ({tipo_nota})...")
                            
                        # Task 4.2: Filtro por Data e Tipo
                        try:
                            # 1. Seleciona o Radio Button (Emitente ou Destinatário)
                            if tipo_nota == "Emitente":
                                page_modulo.locator("#rdEmitente").check(timeout=5000)
                            else:
                                page_modulo.locator("#rdDestinatario").check(timeout=5000)
                                
                            time.sleep(1)
                            
                            # 2. Preenche os campos de Período (Data Inicial e Data Final) usando os IDs corretos
                            page_modulo.locator("#periodoIni").fill("")
                            page_modulo.locator("#periodoIni").type(data_inicial)
                            
                            page_modulo.locator("#periodoFim").fill("")
                            page_modulo.locator("#periodoFim").type(data_final)
                            
                            # 3. Clica no botão "Consultar" pelo ID exato
                            page_modulo.locator("#btnConsultar").click(timeout=5000)
                            
                        except Exception as e:
                            if log_callback: log_callback(f"Aviso: Erro ao preencher o formulário para {tipo_nota}: {e}")

                        if log_callback: log_callback("Aguardando resultados da consulta...")
                        time.sleep(4) # Aguarda processamento da pesquisa
                        
                        # Task 4.3: Tratar mensagem "Consulta não Retornou Registros"
                        try:
                            msg_vazia = page_modulo.get_by_text("Consulta não Retornou Registros")
                            if msg_vazia.is_visible(timeout=3000):
                                if log_callback:
                                    log_callback(f"  -> Nenhuma nota encontrada para {tipo_nota} nesta competência.")
                                return dados
                        except Exception:
                            pass
                            
                        # Task 4.4: Lidar com paginação (150 registros por página) e coletar dados
                        try:
                            # Tenta puxar direto pela API do jqGrid (muito mais rápido e extrai a grade inteira)
                            dados_grid = page_modulo.evaluate('''() => {
                                return $("table.ui-jqgrid-btable").jqGrid("getRowData");
                            }''')
                            if dados_grid and len(dados_grid) > 0:
                                for row in dados_grid:
                                    dados.append(row)
                                if log_callback: log_callback(f"  -> Extraídas {len(dados_grid)} notas usando API do Grid!")
                                return dados
                        except Exception as e:
                            pass
                            
                        # Fallback: Extração ultra-rápida via JavaScript direto no navegador
                        if log_callback: log_callback("Iniciando varredura da grade de resultados...")
                        while True:
                            try:
                                page_modulo.wait_for_selector("table.ui-jqgrid-btable tr.jqgrow", timeout=5000)
                                time.sleep(1) # Aguardar a animação do loader
                                
                                # Extrai toda a tabela de uma vez só executando JS na página (Muito mais rápido!)
                                js_extrair_tabela = """
                                () => {
                                    let colunas = Array.from(document.querySelectorAll("table.ui-jqgrid-htable th"))
                                        .filter(th => th.offsetWidth > 0 || th.offsetHeight > 0 || th.getClientRects().length > 0)
                                        .map((th, i) => {
                                            let text = th.innerText.trim();
                                            return text ? text : "Col_" + i;
                                        });

                                    let linhas = Array.from(document.querySelectorAll("table.ui-jqgrid-btable tr.jqgrow"));
                                    let resultados = linhas.map(linha => {
                                        let celulas = Array.from(linha.querySelectorAll("td")).filter(td => td.style.display !== "none" && (td.offsetWidth > 0 || td.offsetHeight > 0));
                                        let obj = {};
                                        celulas.forEach((td, idx) => {
                                            let nome_col = colunas[idx] || "Coluna_" + (idx+1);
                                            obj[nome_col] = td.innerText.trim();
                                        });
                                        return obj;
                                    });
                                    return resultados;
                                }
                                """
                                
                                dados_pagina = page_modulo.evaluate(js_extrair_tabela)
                                if dados_pagina:
                                    dados.extend(dados_pagina)
                                        
                                # Tenta achar o botão de Próxima Página (span.ui-icon-seek-next)
                                btn_proxima = page_modulo.locator("td#next_jqGridPager, .ui-jqgrid-pager .ui-icon-seek-next").first
                                
                                if btn_proxima.is_visible():
                                    # Verifica se o botão "Próximo" está desativado (ui-state-disabled)
                                    is_disabled = page_modulo.evaluate('''() => {
                                        const btn = document.querySelector(".ui-jqgrid-pager .ui-icon-seek-next");
                                        if(!btn) return true;
                                        const td = btn.closest('td');
                                        return td && td.classList.contains('ui-state-disabled');
                                    }''')
                                    
                                    if is_disabled:
                                        break
                                    else:
                                        btn_proxima.click()
                                        time.sleep(2) # Espera a página carregar
                                else:
                                    break
                            except:
                                break
                                
                        if log_callback: log_callback(f"  -> Extraídas {len(dados)} notas manualmente!")
                        return dados

                    # Coleta Notas Emitidas (Emitente)
                    nfe_emitidas = extrair_tabela_paginada("Emitente")
                    
                    # Coleta Notas Recebidas (Destinatário)
                    nfe_recebidas = extrair_tabela_paginada("Destinatário")
                    
                    # --- FASE 5: Captura CT-e ---
                    cte_emitidos = [{"Aviso": "A Consulta não Retornou Registros"}]
                    cte_recebidos = [{"Aviso": "A Consulta não Retornou Registros"}]
                    
                    try:
                        if log_callback: log_callback("Acessando Menu de Opções > Doc Fiscais > CT-e...")
                        
                        # 1. Clica no "Menu de Opções"
                        menu_opcoes = page_modulo.locator("text=/Menu de Opções/i").first
                        menu_opcoes.wait_for(state="visible", timeout=10000)
                        menu_opcoes.click()
                        time.sleep(1.5)
                        
                        # 2. Hover/Click em "Documentos Fiscais Eletrônicos"
                        doc_fiscais = page_modulo.locator("text=/Documentos Fiscais Eletr/i").first
                        doc_fiscais.hover()
                        time.sleep(1)
                        doc_fiscais.click()
                        time.sleep(1)
                        
                        # 3. Clica em "Conhecimento de Transporte Eletrônico"
                        cte_menu = page_modulo.locator("text=/Conhecimento de Transporte Eletr/i").first
                        cte_menu.click()
                        
                        page_modulo.wait_for_load_state("domcontentloaded")
                        time.sleep(3)
                        
                        def extrair_cte(tipo_cte):
                            dados = []
                            if log_callback: log_callback(f"Iniciando pesquisa de CT-e ({tipo_cte})...")
                            try:
                                if tipo_cte == "Emitente":
                                    radio = page_modulo.locator("#rdEmitente").first
                                    radio.check(timeout=5000)
                                else:
                                    # Para Tomador (A Sefaz copiou a tela e manteve o ID como rdDestinatario)
                                    radio = page_modulo.locator("#rdDestinatario").first
                                    radio.check(timeout=5000)
                                    
                                time.sleep(1)
                                
                                # Preencher Datas.
                                page_modulo.locator("#periodoIni").fill("")
                                page_modulo.locator("#periodoIni").type(data_inicial)
                                page_modulo.locator("#periodoFim").fill("")
                                page_modulo.locator("#periodoFim").type(data_final)
                                        
                                # Remove a tabela velha da tela para ter certeza de que o wait_for_selector vai aguardar a nova consulta
                                page_modulo.evaluate("() => { let tbl = document.querySelector('table.ui-jqgrid-btable'); if (tbl) tbl.remove(); }")
                                
                                # Consultar
                                page_modulo.locator("#btnConsultar, button:has-text('Consultar')").first.click(timeout=5000)
                                
                                # Esperar grade carregar
                                if log_callback: log_callback("Aguardando resultados do CT-e...")
                                try:
                                    page_modulo.wait_for_selector(".ui-jqgrid-btable, .sweet-alert, div.alert, span:has-text('não retornou registros')", timeout=15000)
                                except:
                                    pass
                                    
                                time.sleep(3)
                                
                                # Verificar alerta (ex: sem registros)
                                erro_alerta = page_modulo.locator(".sweet-alert.visible").first
                                if erro_alerta.is_visible():
                                    texto = erro_alerta.inner_text().lower()
                                    if 'não retornou' in texto or 'nenhum registro' in texto:
                                        if log_callback: log_callback(f"  -> Nenhum CT-e localizado para {tipo_cte}.")
                                        btn_ok = erro_alerta.locator("button.confirm").first
                                        if btn_ok.is_visible(): btn_ok.click()
                                        return [{"Aviso": "A Consulta não Retornou Registros"}]
                                
                                js_extrair_tabela = """
                                () => {
                                    let tbl = document.querySelector("table.ui-jqgrid-btable");
                                    if(!tbl) return [];
                                    
                                    let colunas = Array.from(document.querySelectorAll("table.ui-jqgrid-htable th"))
                                        .filter(th => th.offsetWidth > 0 || th.offsetHeight > 0 || th.getClientRects().length > 0)
                                        .map((th, i) => {
                                            let text = th.innerText.trim();
                                            return text ? text : "Col_" + i;
                                        });

                                    let linhas = Array.from(document.querySelectorAll("table.ui-jqgrid-btable tr.jqgrow"));
                                    let resultados = linhas.map(linha => {
                                        let celulas = Array.from(linha.querySelectorAll("td")).filter(td => td.style.display !== "none" && (td.offsetWidth > 0 || td.offsetHeight > 0));
                                        let obj = {};
                                        celulas.forEach((td, idx) => {
                                            let nome_col = colunas[idx] || "Coluna_" + (idx+1);
                                            obj[nome_col] = td.innerText.trim();
                                        });
                                        return obj;
                                    });
                                    return resultados;
                                }
                                """
                                
                                while True:
                                    try:
                                        page_modulo.wait_for_selector("table.ui-jqgrid-btable tr.jqgrow", timeout=5000)
                                        time.sleep(1)
                                        
                                        dados_pagina = page_modulo.evaluate(js_extrair_tabela)
                                        if dados_pagina and len(dados_pagina) > 0:
                                            dados.extend(dados_pagina)
                                                
                                        btn_proxima = page_modulo.locator("td#next_jqGridPager, .ui-jqgrid-pager .ui-icon-seek-next").first
                                        if btn_proxima.is_visible():
                                            is_disabled = page_modulo.evaluate('''() => {
                                                const btn = document.querySelector('td#next_jqGridPager, .ui-jqgrid-pager .ui-icon-seek-next');
                                                if (!btn) return true;
                                                const td = btn.closest('td');
                                                return td && td.classList.contains('ui-state-disabled');
                                            }''')
                                            
                                            if is_disabled:
                                                break
                                            else:
                                                btn_proxima.click()
                                                time.sleep(2)
                                        else:
                                            break
                                    except Exception:
                                        break
                                        
                                if len(dados) == 0:
                                    return [{"Aviso": "A Consulta não Retornou Registros"}]
                                    
                                if log_callback: log_callback(f"  -> Extraídos {len(dados)} CT-e ({tipo_cte})!")
                                return dados

                            except Exception as e:
                                if log_callback: log_callback(f"Aviso: Erro ao pesquisar CT-e para {tipo_cte}: {e}")
                                return [{"Aviso": "A Consulta não Retornou Registros"}]

                        cte_emitidos = extrair_cte("Emitente")
                        cte_recebidos = extrair_cte("Tomador")
                        
                    except Exception as e:
                        if log_callback: log_callback(f"Aviso: Erro ao acessar a área de CT-e: {e}")

                    # Fecha a aba se tiver aberto em uma nova, para não acumular
                    if page_modulo != page:
                        page_modulo.close()
                        page.bring_to_front()

                    # --- Fase 6: Gerar Excel Individual da Empresa ---
                    if log_callback: log_callback(f"Gerando Excel com 4 abas para {nome_empresa}...")
                    
                    try:
                        # Se não for informado um diretório de saída no frontend, salva na mesma pasta do arquivo base
                        out_dir = output_dir if output_dir else os.path.dirname(filepath)
                        os.makedirs(out_dir, exist_ok=True)
                        
                        # Nomenclatura mmaaaa_Sefaz_Empresa.xlsx
                        comp_limpo = comp.replace('/', '') # "04/2026" -> "042026"
                        nome_empresa_limpo = "".join([c for c in nome_empresa if c.isalnum() or c in (' ', '-', '_')]).strip()
                        
                        base_filename = f"{comp_limpo}_Sefaz_{nome_empresa_limpo}.xlsx"
                        output_file = os.path.join(out_dir, base_filename)
                        
                        # Adiciona (1), (2), etc caso arquivo já exista
                        contador = 1
                        while os.path.exists(output_file):
                            output_file = os.path.join(out_dir, f"{comp_limpo}_Sefaz_{nome_empresa_limpo} ({contador}).xlsx")
                            contador += 1
                        
                        # Função auxiliar para limpar e formatar dados numéricos
                        def processar_df(dados):
                            df = pd.DataFrame(dados)
                            if not df.empty:
                                for col in df.columns:
                                    if df[col].dtype == object:
                                        # Remove tags HTML perdidas
                                        df[col] = df[col].astype(str).str.replace(r'<[^>]+>', '', regex=True)
                                        
                                        col_lower = str(col).lower()
                                        if any(x in col_lower for x in ['valor', 'vlr', 'icms', 'bc', 'total', 'peso']):
                                            # Limpa o texto monetário padrão "1.234,56" para "1234.56" que o pandas aceita
                                            # Trata NaN/'None' primeiro para evitar erros
                                            df[col] = df[col].replace('None', '').replace('nan', '')
                                            df[col] = df[col].str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
                                            df[col] = pd.to_numeric(df[col], errors='ignore')
                            return df
                            
                        # Função auxiliar para criar resumo (Totalizadores) no rodapé
                        def adicionar_resumo(df, tipo):
                            if df.empty or 'Aviso' in df.columns:
                                return df
                            try:
                                uf_col = next((c for c in df.columns if 'UF' in str(c)), None)
                                sit_col = next((c for c in df.columns if 'Sit' in str(c)), None)
                                
                                resumo_rows = []
                                resumo_rows.append({c: '' for c in df.columns}) # Espaço em branco
                                
                                # Descobre dinamicamente colunas monetárias para somar
                                colunas_soma = [c for c in df.columns if any(x in str(c).lower() for x in ['total nf', 'cálc', 'c\u00e1lc', 'icms', 'prest', 'bc', 'valor'])]
                                
                                if tipo == "NFe Emitida" and uf_col and sit_col:
                                    # Agrupar por UF (somente Autorizadas 'A')
                                    df_aut = df[df[sit_col] == 'A']
                                    if not df_aut.empty:
                                        for uf, group in df_aut.groupby(uf_col):
                                            row = {c: '' for c in df.columns}
                                            row[df.columns[0]] = len(group)
                                            row[uf_col] = uf
                                            for c in colunas_soma: row[c] = group[c].sum()
                                            resumo_rows.append(row)
                                            
                                    resumo_rows.append({c: '' for c in df.columns})
                                    # Total por Situação
                                    for sit, group in df.groupby(sit_col):
                                        row = {c: '' for c in df.columns}
                                        row[df.columns[0]] = len(group)
                                        row[uf_col] = 'total'
                                        for c in colunas_soma: row[c] = group[c].sum()
                                        row[sit_col] = sit
                                        resumo_rows.append(row)
                                        
                                elif tipo == "NFe Recebida" and uf_col:
                                    # Agrupar por UF (Todas as Situações se houver)
                                    for uf, group in df.groupby(uf_col):
                                        row = {c: '' for c in df.columns}
                                        row[df.columns[0]] = len(group)
                                        row[uf_col] = uf
                                        for c in colunas_soma: row[c] = group[c].sum()
                                        resumo_rows.append(row)
                                        
                                    # Total Geral
                                    row = {c: '' for c in df.columns}
                                    row[df.columns[0]] = len(df)
                                    row[uf_col] = 'total'
                                    for c in colunas_soma: row[c] = df[c].sum()
                                    resumo_rows.append(row)
                                    
                                elif "CTe" in tipo and sit_col:
                                    # Agrupar por Situação
                                    for sit, group in df.groupby(sit_col):
                                        row = {c: '' for c in df.columns}
                                        row[df.columns[0]] = len(group)
                                        for c in colunas_soma: row[c] = group[c].sum()
                                        row[sit_col] = sit
                                        resumo_rows.append(row)
                                        
                                    # Total Geral
                                    row = {c: '' for c in df.columns}
                                    row[df.columns[0]] = len(df)
                                    idx_sit = list(df.columns).index(sit_col)
                                    if idx_sit > 0:
                                        row[df.columns[idx_sit - 1]] = 'total'
                                    for c in colunas_soma: row[c] = df[c].sum()
                                    resumo_rows.append(row)
                                    
                                if len(resumo_rows) > 1:
                                    df_resumo = pd.DataFrame(resumo_rows)
                                    df = pd.concat([df, df_resumo], ignore_index=True)
                            except Exception as e:
                                print(f"Erro ao adicionar resumo: {e}")
                            return df
                            
                        df_nfe_e = adicionar_resumo(processar_df(nfe_emitidas), "NFe Emitida")
                        df_nfe_r = adicionar_resumo(processar_df(nfe_recebidas), "NFe Recebida")
                        df_cte_e = adicionar_resumo(processar_df(cte_emitidos), "CTe Emitido")
                        df_cte_r = adicionar_resumo(processar_df(cte_recebidos), "CTe Recebido")
                        
                        # Grava o arquivo com engine xlsxwriter para permitir formatação fina
                        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                            def salvar_aba(df, nome_aba):
                                df.to_excel(writer, sheet_name=nome_aba, index=False)
                                if df.empty: return
                                
                                workbook = writer.book
                                worksheet = writer.sheets[nome_aba]
                                
                                # Formatadores
                                header_format = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
                                num_format = workbook.add_format({'num_format': '#,##0.00'})
                                text_format = workbook.add_format({'num_format': '@'}) # Text puro (evita E+14)
                                
                                for col_num, value in enumerate(df.columns.values):
                                    # Pinta o cabeçalho
                                    worksheet.write(0, col_num, value, header_format)
                                    
                                    # Calcula largura ideal
                                    try:
                                        column_len = max(df[value].astype(str).map(len).max(), len(str(value))) + 2
                                    except:
                                        column_len = 15
                                    column_len = min(column_len, 40) # Teto máximo de largura
                                    
                                    col_lower = str(value).lower()
                                    if any(x in col_lower for x in ['chave', 'cnpj', 'ie', 'número', 'numero', 'nfe']):
                                        worksheet.set_column(col_num, col_num, column_len, text_format)
                                    elif any(x in col_lower for x in ['valor', 'vlr', 'icms', 'bc', 'total', 'peso']):
                                        worksheet.set_column(col_num, col_num, column_len, num_format)
                                    else:
                                        worksheet.set_column(col_num, col_num, column_len)

                            salvar_aba(df_nfe_e, 'NFe Emitida')
                            salvar_aba(df_nfe_r, 'NFe Recebida')
                            salvar_aba(df_cte_e, 'CTe Emitido')
                            salvar_aba(df_cte_r, 'CTe Recebido')
                            
                        if log_callback: log_callback(f"✅ Arquivo salvo: {output_file}")
                    except Exception as e:
                        if log_callback: log_callback(f"❌ Erro ao gerar Excel para {nome_empresa}: {e}")

                except Exception as e:
                    if log_callback:
                        log_callback(f"❌ Erro ao configurar {nome_empresa}: {str(e)}")
                    # Segue para a próxima empresa mesmo se essa falhar
                    continue

            if log_callback:
                log_callback("==================================================")
                log_callback("Fim do processamento de todas as empresas da lista.")
                log_callback("✅ Processo da Sefaz finalizado com sucesso.")

        except Exception as e:
            if log_callback:
                log_callback(f"Erro crítico na automação Sefaz: {str(e)}")
            raise e

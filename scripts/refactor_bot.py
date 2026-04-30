import os

filepath = r"c:\APP\GestorPrefeitura\src\nfse_bot.py"
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# I will replace the block starting at '# 3. Informar Datas e Filtros no form'
# up to the end of the success logging (line 516 approx) with our new looped version.

old_block = """                # 3. Informar Datas e Filtros no form
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
                if log_callback: log_callback(f"☁️ Solicitando geração do relatório para {cnpj}...")
                
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
                    print(f"Botão de Excel não localizado. Capturando tela de erro para {cnpj}.")
                    page.screenshot(path=os.path.join(errors_dir, f"erro_btn_excel_{cnpj}_{timestamp_log}.png"))
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
                    if log_callback: log_callback(f"📂 {cnpj} não possui notas no período (Zerado).")
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
                    log_exec(cnpj, nome_empresa, competencia_mes_ano, "SUCESSO_VAZIO", "Sem notas fiscais no período. Valores zerados registrados.", index=idx, total=total_targets, csv_path=csv_log_path)
                    continue # Pula o processamento do excel e vai pra rotina do proximo cnpj
                    
                if not download_obj[0]:
                    log_exec(cnpj, nome_empresa, competencia_mes_ano, "ERRO_TIMEOUT", "Timeout aguardando download ou mensagem de erro", index=idx, total=total_targets, csv_path=csv_log_path)
                    raise Exception("Timeout aguardando download ou mensagem de erro da prefeitura.")
                    
                download = download_obj[0]
                
                # Resiliencia de nome gerado
                sugg_ext = download.suggested_filename.split('.')[-1]
                if not sugg_ext: sugg_ext = 'xlsx'
                
                # Previne que barras ou caracteres ilegais no nome da empresa (ex: S/S) criem pastas indesejadas
                nome_seguro = nome_empresa.replace('/', '-').replace('\\', '-').strip()
                if not nome_seguro:
                    nome_seguro = 'empresa'
                    
                # Diretorio destino customizado ou padrao
                out_dir = output_dir if output_dir else LIVROS_DIR
                
                filepath = os.path.join(out_dir, f"{competencia_mes_ano.replace('/','')}_{nome_seguro}.{sugg_ext}")
                if os.path.exists(filepath):
                    agora = datetime.now().strftime("%H%M%S")
                    filepath = os.path.join(out_dir, f"{competencia_mes_ano.replace('/','')}_{nome_seguro}_{agora}.{sugg_ext}")
                    
                download.save_as(filepath)
                print(f"Download salvo em: {filepath}")
                if log_callback: log_callback(f"💾 Arquivo salvo: {os.path.basename(filepath)}")
                
                # 5. Processamento Local (Pandas)
                resumo = process_livro_fiscal(filepath, target, competencia_mes_ano)
                if resumo:
                    # Remove linha de TOTAL anterior se existir para nao acumular no DF em memoria
                    master_df = master_df[master_df["CNPJ"] != "TOTAL"]
                    
                    master_df = pd.concat([master_df, pd.DataFrame([resumo])], ignore_index=True)
                    
                    # Adiciona Linha de Totais ao final
                    numeric_cols = [
                        "Qtd_Notas_Ativas", "Qtd_Notas_Canceladas", "Faturamento_Bruto", 
                        "ISS_Proprio", "ISS_Retido", "Valor_INSS", "Valor_IR", 
                        "Valor_COFINS", "Valor_CSLL", "Valor_PIS", "Valor_Liquido"
                    ]
                    # Garante que as colunas sejam numericas antes de somar e arredonda para 2 decimais (padrao moeda)
                    for col in numeric_cols:
                        master_df[col] = pd.to_numeric(master_df[col], errors='coerce').fillna(0).round(2)
                        
                    totals = master_df[numeric_cols].sum().round(2)
                    row_total = {col: totals[col] for col in numeric_cols}
                    row_total.update({"Mes_Competencia": competencia_mes_ano, "CNPJ": "TOTAL", "Nome": "SOMA TOTAL DOS PROCESSADOS"})
                    
                    final_df = pd.concat([master_df, pd.DataFrame([row_total])], ignore_index=True)
                    save_master_df(final_df, config_file_path)
                    
                    print(f"Planilha de consolidação atualizada no arquivo de alvos para {cnpj}.")
                    if log_callback: log_callback(f"✔️ Dados de {cnpj} consolidados no arquivo mestre.")
                    log_exec(cnpj, nome_empresa, competencia_mes_ano, "SUCESSO", f"Arquivo {sugg_ext.upper()} consolidado na aba interna", index=idx, total=total_targets, csv_path=csv_log_path)"""

new_block = """                # NOVO LOOP DE MÚLTIPLOS AGRUPAMENTOS: Prestados e Tomados
                for tipo_agrupamento, tipo_sigla in [('Serviços Prestados', 'ptd'), ('Serviços Tomados', 'tmd')]:
                    if log_callback: log_callback(f"Extraindo {tipo_agrupamento}...")
                    try:
                        # 3. Informar o Agrupamento no form do Primefaces
                        try:
                            lbl_agrup = page.locator('label', has_text='Agrupamento').first
                            if lbl_agrup.is_visible():
                                lbl_agrup.locator('xpath=..').locator('.ui-selectonemenu-trigger').click(force=True)
                                time.sleep(1)
                                item_tipo = page.locator(f'li.ui-selectonemenu-item:has-text("{tipo_agrupamento}")').first
                                if item_tipo.is_visible():
                                    item_tipo.click(force=True)
                                else:
                                    js_click_text('li.ui-selectonemenu-item', tipo_agrupamento)
                                time.sleep(1)
                        except Exception as e:
                            print(f"Nota: não foi possível mudar 'Agrupamento' para {tipo_agrupamento}. Erro: {e}")

                        # 4. Informar Datas e Filtros no form
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
                                
                        # 5. Tentar encontrar a Saída "Excel" antes de clicar no export / gerar
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

                        # 6. Baixar Excel
                        print(f"Aguardando geracao do excel ou mensagem de erro (vazio) para {tipo_sigla}...")
                        
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
                            print(f"Botão de Excel não localizado para {tipo_sigla}. Capturando tela.")
                            page.screenshot(path=os.path.join(errors_dir, f"erro_btn_excel_{cnpj}_{tipo_sigla}_{timestamp_log}.png"))
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
                                    
                            page.wait_for_timeout(1000)
                            timeout_count += 1
                            
                        # Limpa o listener
                        page.remove_listener("download", handle_download)
                        
                        if has_error:
                            print(f"Empresa {cnpj} não possui notas {tipo_agrupamento} na competência {competencia_mes_ano}.")
                            if log_callback: log_callback(f"📂 {cnpj} não possui notas {tipo_sigla} no período (Zerado).")
                            log_exec(cnpj, nome_empresa, competencia_mes_ano, "VAZIO", f"Sem notas de {tipo_agrupamento}.", index=idx, total=total_targets, csv_path=csv_log_path)
                            continue 
                            
                        if not download_obj[0]:
                            log_exec(cnpj, nome_empresa, competencia_mes_ano, "ERRO_TIMEOUT", f"Timeout aguardando download {tipo_sigla}", index=idx, total=total_targets, csv_path=csv_log_path)
                            print(f"Timeout no {tipo_sigla} da empresa {cnpj}")
                            continue
                            
                        download = download_obj[0]
                        sugg_ext = download.suggested_filename.split('.')[-1]
                        if not sugg_ext: sugg_ext = 'xlsx'
                        
                        nome_seguro = nome_empresa.replace('/', '-').replace('\\\\', '-').strip()
                        if not nome_seguro:
                            nome_seguro = 'empresa'
                            
                        out_dir = output_dir if output_dir else LIVROS_DIR
                        
                        # NOVO PADRÃO DE NOME + RESILIENCIA COM (1), (2), etc
                        base_filename = f"{competencia_mes_ano.replace('/','')}-{tipo_sigla}-{nome_seguro}"
                        filepath = os.path.join(out_dir, f"{base_filename}.{sugg_ext}")
                        
                        counter = 1
                        while os.path.exists(filepath):
                            filepath = os.path.join(out_dir, f"{base_filename}({counter}).{sugg_ext}")
                            counter += 1
                            
                        download.save_as(filepath)
                        print(f"Download salvo em: {filepath}")
                        if log_callback: log_callback(f"💾 Arquivo {tipo_sigla} salvo: {os.path.basename(filepath)}")
                        
                        # 7. Processamento e Totais Ocultos
                        sucesso_totais = process_livro_fiscal(filepath, target, competencia_mes_ano)
                        if sucesso_totais:
                            log_exec(cnpj, nome_empresa, competencia_mes_ano, "SUCESSO", f"Arquivo {tipo_sigla.upper()} salvo com Totais", index=idx, total=total_targets, csv_path=csv_log_path)
                        else:
                            log_exec(cnpj, nome_empresa, competencia_mes_ano, "SUCESSO_PARCIAL", f"Arquivo {tipo_sigla.upper()} salvo, mas falha ao injetar Totais Localmente", index=idx, total=total_targets, csv_path=csv_log_path)
                            
                    except Exception as loop_e:
                        print(f"Erro processando {tipo_sigla} de {cnpj}: {loop_e}")
                        log_exec(cnpj, nome_empresa, competencia_mes_ano, "ERRO_TIPO", f"Falha no tipo {tipo_sigla}: {loop_e}", index=idx, total=total_targets, csv_path=csv_log_path)"""

if old_block in content:
    content = content.replace(old_block, new_block)
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)
    print("Sucesso: Substituído.")
else:
    print("O bloco original nao foi encontrado de forma exata na string!")

import os
import glob
import pandas as pd
import sys

# Setup caminhos
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PLAN_DIR = os.path.join(BASE_DIR, "planejamento")
LIVROS_DIR = os.path.join(PLAN_DIR, "livros")
CONFIG_FILE = os.path.join(PLAN_DIR, "Relação de empresas.xlsx")

def clean_curr(series):
    if series.dtype == 'object':
        return pd.to_numeric(series.astype(str).str.replace(r'[R$\s]', '', regex=True).str.replace('.', '', regex=False).str.replace(',', '.'), errors='coerce').fillna(0)
    return pd.to_numeric(series, errors='coerce').fillna(0)

def processar_arquivos_existentes():
    if not os.path.exists(LIVROS_DIR):
        print(f"Diretório não encontrado: {LIVROS_DIR}")
        return

    arquivos = glob.glob(os.path.join(LIVROS_DIR, "*.xlsx"))
    
    # Vamos carregar os alvos só pra pegar CNPJs
    df_raw = pd.read_excel(CONFIG_FILE, header=None)
    header_idx = 0
    for i in range(min(10, len(df_raw))):
        row_vals = [str(v).lower() for v in df_raw.iloc[i].values]
        if any('cnpj' in v for v in row_vals):
            header_idx = i
            break
            
    df_alvos = pd.read_excel(CONFIG_FILE, header=header_idx)
    df_alvos = df_alvos.dropna(how='all')
    
    col_nome, col_cnpj = df_alvos.columns[1], df_alvos.columns[2]
    
    targets = {}
    for _, row in df_alvos.iterrows():
        nome = str(row[col_nome]).strip() if pd.notna(row[col_nome]) else ""
        cnpj_raw = str(row[col_cnpj]).strip() if pd.notna(row[col_cnpj]) else ""
        cnpj_clean = "".join(filter(str.isdigit, cnpj_raw))
        if len(cnpj_clean) >= 11:
            nome_seguro = nome.replace('/', '-').replace('\\', '-').strip()
            targets[nome_seguro.lower()] = {'CNPJ': cnpj_clean, 'Nome': nome}

    resultados = []
    
    print(f"Encontrados {len(arquivos)} arquivos Excel para consolidar.")
    
    for arquivo in arquivos:
        nome_arquivo = os.path.basename(arquivo)
        # Tenta achar de qual empresa é o arquivo baseado no nome
        alvo_atual = None
        for key_nome, dados in targets.items():
            if key_nome in nome_arquivo.lower():
                alvo_atual = dados
                break
                
        if not alvo_atual:
            print(f"? Ignorando {nome_arquivo}: não encontrou o nome correspondente na lista.")
            continue
            
        try:
            df = pd.read_excel(arquivo, header=5)
        except Exception as e:
            print(f"Erro lendo {nome_arquivo}: {e}")
            continue
            
        df_ativas = df[df['SITUACAO'].str.lower() == 'ativa'].copy()
        df_canceladas = df[df['SITUACAO'].str.lower() == 'cancelada'].copy()
        
        for col in df_ativas.columns:
            if 'valor' in str(col).lower() or 'retido' in str(col).lower():
                df_ativas[col] = clean_curr(df_ativas[col])
                
        fat_bruto = df_ativas['VALOR NF'].sum() if 'VALOR NF' in df_ativas.columns else 0
        iss = df_ativas['VALOR ISS'].sum() if 'VALOR ISS' in df_ativas.columns else 0
        inss = df_ativas['VALOR INSS'].sum() if 'VALOR INSS' in df_ativas.columns else 0
        ir = df_ativas['VALOR IR'].sum() if 'VALOR IR' in df_ativas.columns else 0
        cof = df_ativas['VALOR COFINS'].sum() if 'VALOR COFINS' in df_ativas.columns else 0
        csl = df_ativas['VALOR CSLL'].sum() if 'VALOR CSLL' in df_ativas.columns else 0
        pis = df_ativas['VALOR PIS'].sum() if 'VALOR PIS' in df_ativas.columns else 0
        liq = df_ativas['VALOR LÍQUIDO'].sum() if 'VALOR LÍQUIDO' in df_ativas.columns else 0
        
        resumo = {
            "Mes_Competencia": nome_arquivo.split('_')[0][:2] + "/" + nome_arquivo.split('_')[0][2:],
            "CNPJ": alvo_atual['CNPJ'],
            "Nome": alvo_atual['Nome'],
            "Qtd_Notas_Ativas": len(df_ativas),
            "Qtd_Notas_Canceladas": len(df_canceladas),
            "Faturamento_Bruto": fat_bruto,
            "ISS_Proprio": 0,
            "ISS_Retido": 0,
            "Valor_INSS": inss,  
            "Valor_IR": ir,
            "Valor_COFINS": cof,
            "Valor_CSLL": csl,
            "Valor_PIS": pis,
            "Valor_Liquido": liq
        }
        
        col_retido = next((c for c in df_ativas.columns if 'retido' in str(c).lower()), None)
        if col_retido:
            resumo["ISS_Retido"] = df_ativas[col_retido].sum()
            resumo["ISS_Proprio"] = iss - resumo["ISS_Retido"]
        else:
            resumo["ISS_Proprio"] = iss
            
        resultados.append(resumo)
        
    if not resultados:
        print("Nenhum dado consolidado!")
        return
        
    master_df = pd.DataFrame(resultados)
    
    # Totais
    numeric_cols = [
        "Qtd_Notas_Ativas", "Qtd_Notas_Canceladas", "Faturamento_Bruto", 
        "ISS_Proprio", "ISS_Retido", "Valor_INSS", "Valor_IR", 
        "Valor_COFINS", "Valor_CSLL", "Valor_PIS", "Valor_Liquido"
    ]
    for col in numeric_cols:
        master_df[col] = pd.to_numeric(master_df[col], errors='coerce').fillna(0)
        
    totals = master_df[numeric_cols].sum()
    row_total = {col: totals[col] for col in numeric_cols}
    row_total.update({"Mes_Competencia": master_df.iloc[0]["Mes_Competencia"], "CNPJ": "TOTAL", "Nome": "SOMA TOTAL DOS PROCESSADOS"})
    
    final_df = pd.concat([master_df, pd.DataFrame([row_total])], ignore_index=True)
    
    print("\nAMOSTRA DA CONSOLIDAÇÃO:")
    print(final_df[["Nome", "Faturamento_Bruto", "Qtd_Notas_Ativas"]].head())
    
    print("\nSalvando no arquivo de relação de empresas...")
    try:
        with pd.ExcelWriter(CONFIG_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            final_df.to_excel(writer, sheet_name="Consolidação_Avulsa", index=False)
        print("Salvo com sucesso na aba 'Consolidação_Avulsa'!")
    except Exception as e:
        print(f"Erro ao salvar: {e}")

if __name__ == "__main__":
    processar_arquivos_existentes()

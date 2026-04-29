import os
import sys
import re
import pandas as pd
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog
from PIL import Image
import pytesseract
import io
import eel
from ofxparse import OfxParser
# REMOVIDO o "import codecs", não precisamos mais!
import tempfile # NOVO: Para salvar o Excel temporariamente

# ==============================================================================
# CONFIGURAÇÕES E INICIALIZAÇÃO DO EEL
# ==============================================================================
# Inicia a pasta web onde está o HTML
eel.init('web')

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CAMINHO_TESSERACT_LOCAL = os.path.join(BASE_DIR, "dependencias", "Tesseract-OCR", "tesseract.exe")
CAMINHO_TESSERACT_SISTEMA = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

if os.path.exists(CAMINHO_TESSERACT_LOCAL):
    pytesseract.pytesseract.tesseract_cmd = CAMINHO_TESSERACT_LOCAL
elif os.path.exists(CAMINHO_TESSERACT_SISTEMA):
    pytesseract.pytesseract.tesseract_cmd = CAMINHO_TESSERACT_SISTEMA

# ==============================================================================
# FUNÇÕES DE UI NATIVA EXPOSTAS PARA O JAVASCRIPT
# ==============================================================================
def open_file_dialog(title, filetypes, save=False, multiple=False):
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True) 
    
    if save:
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=f"Relatorio_{datetime.now().strftime('%d-%m-%Y')}.xlsx", title=title, filetypes=filetypes)
        res = filename
    elif multiple:
        # ATIVAMOS O MODO DE SELEÇÃO MÚLTIPLA NO WINDOWS!
        filenames = filedialog.askopenfilenames(title=title, filetypes=filetypes)
        res = list(filenames)
    else:
        filename = filedialog.askopenfilename(title=title, filetypes=filetypes)
        res = filename
        
    root.destroy()
    return res

@eel.expose
def escolher_planilha_py():
    return open_file_dialog("Selecione a Planilha de Fluxo Master", [("Excel/CSV", "*.xlsx *.xls *.csv")])

@eel.expose
def escolher_extratos_multiplos_py():
    # Repare no multiple=True, permitindo a Carol arrastar o mouse sobre os 6 OFXs
    return open_file_dialog("Selecione os Extratos (Pode selecionar vários de uma vez)", [("Extratos Bancários", "*.ofx *.ofc *.txt *.pdf *.xlsx *.xls *.csv")], multiple=True)

@eel.expose
def escolher_onde_salvar_py():
    return open_file_dialog("Salvar Relatório", [("Excel", "*.xlsx")], save=True)

# ==============================================================================
# MOTOR DE CONCILIAÇÃO
# ==============================================================================
# (MANTENHA AS FUNÇÕES excel_date_to_datetime E limpar_valor_inteligente INTACTAS AQUI)
def excel_date_to_datetime(serial):
    try:
        if pd.isna(serial): return None
        if isinstance(serial, (int, float)): return datetime(1899, 12, 30) + timedelta(days=float(serial))
        val_str = str(serial).strip()
        for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d/%m/%Y', '%d/%m/%y'):
            try: return datetime.strptime(val_str.split()[0], fmt)
            except: pass
        return pd.to_datetime(serial)
    except: return None

def limpar_valor_inteligente(val):
    try:
        s_val = str(val).strip().replace('R$', '').replace(' ', '')
        if '.' in s_val and ',' in s_val: s_val = s_val.replace('.', '').replace(',', '.')
        elif ',' in s_val: s_val = s_val.replace(',', '.')
        elif '.' in s_val and s_val.count('.') > 1: s_val = s_val.replace('.', '')
        return float(s_val)
    except: return 0.0

def detectar_colunas(df):
    # AGORA O ROBÔ TAMBÉM CAÇA A COLUNA "BANCO" OU "CONTA"
    col_data, col_valor, col_banco, start_row = None, None, None, 0
    for i, row in df.head(20).iterrows():
        row_txt = [str(x).upper().strip() for x in row.values]
        if "DATA" in row_txt and "VALOR" in row_txt:
            for idx, val in enumerate(row_txt):
                if "DATA" in val: col_data = idx
                if "VALOR" in val: col_valor = idx
                if "BANCO" in val or "CONTA" in val: col_banco = idx # Capturamos o Banco!
            return col_data, col_valor, col_banco, i + 1
    
    amostra = df.iloc[5:15] if len(df) > 15 else df
    melhor_score_data, melhor_score_valor = -1, -1
    for col in df.columns:
        is_date = sum(1 for val in amostra[col] if excel_date_to_datetime(val) is not None)
        if is_date > melhor_score_data: melhor_score_data, col_data = is_date, col
    for col in df.columns:
        if col == col_data: continue
        is_num = sum(1 for val in amostra[col] if isinstance(limpar_valor_inteligente(val), float) and limpar_valor_inteligente(val) != 0)
        if is_num >= melhor_score_valor: melhor_score_valor, col_valor = is_num, col
            
    return col_data, col_valor, col_banco, 0

@eel.expose
def iniciar_conciliacao_py(caminho_planilha, lista_caminhos_extratos):
    try:
        # --- 1. LER PLANILHA DE FLUXO MASTER ---
        if caminho_planilha.endswith('.csv'):
            try: df_raw = pd.read_csv(caminho_planilha, sep=';', header=None)
            except: df_raw = pd.read_csv(caminho_planilha, sep=',', header=None)
        else:
            xls = pd.ExcelFile(caminho_planilha)
            df_raw = pd.read_excel(xls, sheet_name=0, header=None)

        col_data_idx, col_valor_idx, col_banco_idx, start_row = detectar_colunas(df_raw)
        dados_planilha = []
        for i, row in df_raw.iterrows():
            if i < start_row: continue
            try:
                dt = excel_date_to_datetime(row[col_data_idx])
                if dt is None: continue
                val = limpar_valor_inteligente(row[col_valor_idx])
                banco_planilha = ""
                if col_banco_idx is not None and pd.notna(row[col_banco_idx]):
                    banco_planilha = str(row[col_banco_idx]).strip()
                
                # Salva a data completa (YYYY-MM-DD)
                dia_str = dt.strftime('%Y-%m-%d')
                dados_planilha.append({'Dia': dia_str, 'Valor': abs(val), 'Banco': banco_planilha})
            except: continue

        if not dados_planilha:
             return {"status": "erro", "erro": "Planilha de fluxo vazia ou sem datas válidas."}

        # =======================================================
        # O PULO DO GATO (V2): Filtro por Intervalo (Range)
        # Identifica a menor e a maior data que a Carol colocou na planilha
        # Ex: Vai pegar "2026-04-01" até "2026-04-13"
        # =======================================================
        datas_planilha_lista = [d['Dia'] for d in dados_planilha]
        data_inicio_fluxo = min(datas_planilha_lista)
        data_fim_fluxo = max(datas_planilha_lista)

        df_planilha = pd.DataFrame(dados_planilha).groupby('Dia')['Valor'].sum().reset_index()
        df_planilha.rename(columns={'Valor': 'Total_Planilha'}, inplace=True)
        df_planilha['Total_Planilha'] = df_planilha['Total_Planilha'].round(2)

        # --- 2. LER O SUPER BOLO DE EXTRATOS BANCÁRIOS ---
        dados_extrato = []
        
        for caminho_extrato in lista_caminhos_extratos:
            nome_arquivo = os.path.basename(caminho_extrato) 

            if caminho_extrato.lower().endswith(('.ofx', '.ofc')):
                try:
                    with open(caminho_extrato, 'r', encoding='latin-1', errors='ignore') as fileobj:
                        ofx = OfxParser.parse(fileobj)
                        for account in ofx.accounts:
                            for transaction in account.statement.transactions:
                                
                                dia_str = transaction.date.strftime('%Y-%m-%d')
                                
                                # SE O DIA DO OFX ESTIVER FORA DO PERÍODO DO FLUXO (Ex: Dia 14 e 15), ELE IGNORA!
                                if not (data_inicio_fluxo <= dia_str <= data_fim_fluxo):
                                    continue
                                
                                val = abs(float(transaction.amount))
                                historico = transaction.memo if transaction.memo else transaction.payee
                                
                                dados_extrato.append({
                                    'Dia': dia_str, 
                                    'Valor': val, 
                                    'Histórico': str(historico).strip(),
                                    'Origem': nome_arquivo
                                })
                except Exception as e:
                    return {"status": "erro", "erro": f"Erro no arquivo {nome_arquivo}: {str(e)}"}

        if not dados_extrato:
            return {"status": "erro", "erro": f"Os OFXs não possuem dados no período analisado: de {data_inicio_fluxo} até {data_fim_fluxo}."}

        df_ext_resumo = pd.DataFrame(dados_extrato).copy()
        df_ext_resumo['Valor'] = -df_ext_resumo['Valor'] 
        df_extrato_agrupado = df_ext_resumo.groupby('Dia')['Valor'].sum().reset_index()
        df_extrato_agrupado.rename(columns={'Valor': 'Total_Extrato'}, inplace=True)
        df_extrato_agrupado['Total_Extrato'] = df_extrato_agrupado['Total_Extrato'].round(2)

        # --- 3. MODO AUDITORIA ---
        auditoria = []
        dias_unicos = sorted(list(set(datas_planilha_lista + [d['Dia'] for d in dados_extrato])))

        for dia in dias_unicos:
            itens_plan = [d for d in dados_planilha if d['Dia'] == dia]
            itens_ext = [d for d in dados_extrato if d['Dia'] == dia]

            itens_ext_sobra = itens_ext.copy()
            itens_plan_sobra = []

            for plan_dict in itens_plan:
                encontrou = False
                for ext_dict in itens_ext_sobra:
                    if plan_dict['Valor'] == ext_dict['Valor']:
                        itens_ext_sobra.remove(ext_dict)
                        encontrou = True
                        break 
                if not encontrou:
                    itens_plan_sobra.append(plan_dict)

            for p_sobra in itens_plan_sobra:
                auditoria.append({'Dia': dia, 'Erro': '⚠️ BANCO - NÃO ENCONTRADO', 'Valor (R$)': p_sobra['Valor'], 'Detalhe': 'Faltou cair na conta', 'Origem': p_sobra['Banco']})
            for e_sobra in itens_ext_sobra:
                auditoria.append({'Dia': dia, 'Erro': '🚨 FLUXO - NÃO ENCONTRADO', 'Valor (R$)': e_sobra['Valor'], 'Detalhe': e_sobra['Histórico'], 'Origem': e_sobra['Origem']})

        df_auditoria = pd.DataFrame(auditoria)
        
        # --- 4. GERAÇÃO DO DADO PARA O DASHBOARD ---
        df_final = pd.merge(df_planilha, df_extrato_agrupado, on='Dia', how='outer').fillna(0)
        df_final['Diferença'] = (df_final['Total_Planilha'] + df_final['Total_Extrato']).round(2)
        df_final['Status'] = df_final['Diferença'].apply(lambda x: "OK" if abs(x) < 0.05 else "DIVERGENTE")
        df_final = df_final.sort_values('Dia')
        divergencias = len(df_final[df_final['Status'] == 'DIVERGENTE'])

        # FILTRO DE INTELIGÊNCIA DA AUDITORIA
        dias_perfeitos = df_final[df_final['Status'] == 'OK']['Dia'].tolist()
        if not df_auditoria.empty and 'Dia' in df_auditoria.columns:
            df_auditoria = df_auditoria[~df_auditoria['Dia'].isin(dias_perfeitos)]
        
        # Transforma a Data pro visual do Brasil na tela
        df_final['Dia'] = pd.to_datetime(df_final['Dia']).dt.strftime('%d/%m/%Y')
        dados_resumo = df_final.to_dict('records')

        if df_auditoria.empty:
            df_auditoria = pd.DataFrame([{"Mensagem": "Parabéns! Nenhuma divergência encontrada nas empresas."}])
            dados_auditoria_front = []
        else:
            df_auditoria = df_auditoria.sort_values('Dia') # Garante que a auditoria também fique na ordem
            df_auditoria['Dia'] = pd.to_datetime(df_auditoria['Dia']).dt.strftime('%d/%m/%Y')
            dados_auditoria_front = df_auditoria.to_dict('records')

        # --- 5. SALVAR ARQUIVO COM ABAS INDIVIDUAIS ---
        import tempfile
        temp_dir = tempfile.gettempdir()
        temp_excel_path = os.path.join(temp_dir, "Relatorio_Temporario_Nascimento.xlsx")
        
        # AGORA O EXTRATO CRU É ORDENADO DE FORMA CRONOLÓGICA!
        df_extrato_cru = pd.DataFrame(dados_extrato).sort_values('Dia').rename(columns={'Valor': 'Valor (R$)'})
        
        with pd.ExcelWriter(temp_excel_path, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='RESUMO GERAL', index=False)
            
            df_ext_completo = pd.DataFrame(dados_extrato)
            df_plan_completa = pd.DataFrame(dados_planilha) 
            
            if not df_ext_completo.empty and 'Origem' in df_ext_completo.columns:
                origens_unicas = df_ext_completo['Origem'].unique()
                for origem in origens_unicas:
                    nome_limpo = origem.replace('.ofx', '').replace('.OFX', '').strip()
                    df_banco_especifico = df_ext_completo[df_ext_completo['Origem'] == origem].copy()
                    df_banco_especifico['Valor'] = -df_banco_especifico['Valor'] 
                    resumo_banco_ext = df_banco_especifico.groupby('Dia')['Valor'].sum().reset_index()
                    resumo_banco_ext.rename(columns={'Valor': 'Total_Extrato'}, inplace=True)
                    
                    if not df_plan_completa.empty and 'Banco' in df_plan_completa.columns:
                        df_plan_banco = df_plan_completa[df_plan_completa['Banco'].astype(str).str.contains(nome_limpo, case=False, na=False)].copy()
                    else:
                        df_plan_banco = pd.DataFrame(columns=['Dia', 'Valor'])
                        
                    resumo_plan_banco = df_plan_banco.groupby('Dia')['Valor'].sum().reset_index()
                    resumo_plan_banco.rename(columns={'Valor': 'Total_Planilha'}, inplace=True)
                    
                    df_resumo_final = pd.merge(resumo_plan_banco, resumo_banco_ext, on='Dia', how='outer').fillna(0)
                    if 'Total_Planilha' not in df_resumo_final.columns: df_resumo_final['Total_Planilha'] = 0.0
                    if 'Total_Extrato' not in df_resumo_final.columns: df_resumo_final['Total_Extrato'] = 0.0
                    
                    df_resumo_final['Diferença'] = (df_resumo_final['Total_Planilha'] + df_resumo_final['Total_Extrato']).round(2)
                    df_resumo_final['Status'] = df_resumo_final['Diferença'].apply(lambda x: "OK" if abs(x) < 0.05 else "DIVERGENTE")
                    
                    # Ordena também a aba individual do banco por Dia!
                    df_resumo_final = df_resumo_final.sort_values('Dia')
                    
                    df_resumo_final['Dia'] = pd.to_datetime(df_resumo_final['Dia']).dt.strftime('%d/%m/%Y')
                    nome_aba = f'Resumo {nome_limpo}'[:31]
                    df_resumo_final.to_excel(writer, sheet_name=nome_aba, index=False)

            if not df_auditoria.empty:
                df_auditoria.to_excel(writer, sheet_name='AUDITORIA', index=False)
            
            if not df_extrato_cru.empty:
                df_extrato_cru['Dia'] = pd.to_datetime(df_extrato_cru['Dia']).dt.strftime('%d/%m/%Y')
                df_extrato_cru.to_excel(writer, sheet_name='EXTRATO CONSOLIDADO', index=False)
            
        return {
            "status": "sucesso", 
            "divergencias": divergencias,
            "arquivo_temporario": temp_excel_path, 
            "dados_auditoria": dados_auditoria_front, 
            "dados_resumo": dados_resumo 
        }

    except Exception as e:
        import traceback
        traceback.print_exc()
        return {"status": "erro", "erro": str(e)}
    
# === NOVA FUNÇÃO: PARA BAIXAR O EXCEL SÓ SE O USUÁRIO QUISER ===
import shutil

@eel.expose
def salvar_excel_dashboard_py(caminho_temporario):
    try:
        # Abre a janela pra Carol escolher onde salvar
        caminho_destino = open_file_dialog("Salvar Relatório de Conciliação", [("Excel", "*.xlsx")], save=True)
        if caminho_destino:
            # Copia o arquivo oculto para onde ela escolheu
            shutil.copy2(caminho_temporario, caminho_destino)
            return True
        return False # Ela cancelou
    except:
        return False
       
# ==============================================================================
# START DO APLICATIVO
# ==============================================================================
def on_close(page, sockets):
    """Função ativada quando o usuário clica no 'X' da janela."""
    print("Janela fechada. Liberando a porta de rede imediatamente...")
    import os
    os._exit(0) # Mata o processo instantaneamente

if __name__ == "__main__":
    # Inicia o EEL sem bloquear a thread (block=False)
    eel.start('index.html', size=(650, 650), position=(400, 200), close_callback=on_close, block=False)
    
    # Loop infinito que mantém o sistema vivo, mas permite o encerramento rápido
    while True:
        eel.sleep(1.0)
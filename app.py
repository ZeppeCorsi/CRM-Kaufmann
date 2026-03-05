import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, date, timedelta, time
import calendar
import urllib.parse

# --- CONFIGURAÇÃO GOOGLE SHEETS ---
def conectar_google_sheets():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    # No deploy, certifique-se que o nome no Secrets é exatamente 'gcp_service_account'
    creds_dict = st.secrets["gcp_service_account"]
    
    # IMPORTANTE: Corrige o problema comum de escape da private_key
    if "private_key" in creds_dict:
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    # Verifique se o nome da planilha é exatamente Panorama (case sensitive)
    return client.open("Panorama")

def carregar_aba_gs(nome_aba):
    try:
        sh = conectar_google_sheets()
        worksheet = sh.worksheet(nome_aba)
        data = worksheet.get_all_records()
        
        if not data: # Se a aba estiver vazia
            return pd.DataFrame()
            
        df = pd.DataFrame(data)
        # Limpa e padroniza as colunas
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        if nome_aba == "Agendamentos" and "DATA" in df.columns:
            # Garante que a data seja lida corretamente
            df['DATA_DT'] = pd.to_datetime(df['DATA'], errors='coerce', dayfirst=True).dt.date
            if "HORA" in df.columns:
                df = df.sort_values(by=["DATA_DT", "HORA"])
        return df
    except Exception as e:
        st.error(f"Erro ao carregar aba {nome_aba}: {e}")
        return pd.DataFrame()

def salvar_agendamento_gs(novo_df):
    try:
        sh = conectar_google_sheets()
        worksheet = sh.worksheet("Agendamentos")
        
        # Converte datas e horários para string antes de enviar para o Sheets
        for col in novo_df.columns:
            if novo_df[col].dtype == 'object' or 'date' in str(novo_df[col].dtype):
                novo_df[col] = novo_df[col].astype(str)
        
        valores = novo_df.values.tolist()
        worksheet.append_rows(valores)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Erro ao salvar: {e}")
        return False

# --- FUNÇÃO ATUALIZAR VISITA (PARA O BOTÃO FINALIZAR) ---
def atualizar_visita_gs(indice_original, novo_orc_resp, data_follow, novos_detalhes):
    try:
        sh = conectar_google_sheets()
        worksheet = sh.worksheet("Agendamentos")
        
        # No Google Sheets, a linha 1 é cabeçalho, então o índice do pandas (0) é a linha 2.
        linha_sheets = int(indice_original) + 2
        
        # Atualiza as colunas específicas (Certifique-se da ordem das colunas na sua planilha)
        # Exemplo: Coluna G = REALIZADA, L = DATA FOLLOW, O = NOVO ORCAMENTO
        # Você pode precisar ajustar a letra da coluna abaixo conforme sua planilha real
        worksheet.update_acell(f'G{linha_sheets}', "SIM") 
        worksheet.update_acell(f'L{linha_sheets}', data_follow.strftime("%d/%m/%Y"))
        worksheet.update_acell(f'O{linha_sheets}', novo_orc_resp)
        
        # Detalhes (Coluna H ou conforme sua estrutura)
        # worksheet.update_acell(f'H{linha_sheets}', novos_detalhes)
        
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Erro ao atualizar planilha: {e}")
        return False
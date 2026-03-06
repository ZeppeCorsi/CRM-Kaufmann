import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, date, timedelta
import calendar
import time as t_module

# --- 1. TRAVA DE SEGURANÇA ---
if 'logged_in' not in st.session_state or not st.session_state.logged_in:
    st.error("🚨 Acesso negado. Por favor, faça login na página inicial.")
    st.stop()

# --- 2. FUNÇÕES DE CONEXÃO E DADOS ---
def conectar_google_sheets():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds_info = st.secrets["gcp_service_account"].to_dict()
    if "private_key" in creds_info:
        pk = creds_info["private_key"].strip().strip('"').strip("'").replace("\\n", "\n")
        creds_info["private_key"] = pk
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_info, scope)
    client = gspread.authorize(creds)
    ID_PLANILHA = "1FI41GZwLTglXT4SAXIEyY53AXuheQg7gb_3pz9pWer0"
    return client.open_by_key(ID_PLANILHA)

def carregar_dados_completos():
    try:
        sh = conectar_google_sheets()
        
        # Carrega Agendamentos
        ws_ag = sh.worksheet("Agendamentos")
        df_ag = pd.DataFrame(ws_ag.get_all_records())
        df_ag.columns = [str(c).strip().upper() for c in df_ag.columns]
        
        # Carrega Endereços (Para_Agendar)
        ws_cli = sh.worksheet("Para_Agendar")
        df_cli = pd.DataFrame(ws_cli.get_all_records())
        df_cli.columns = [str(c).strip().upper() for c in df_cli.columns]
        
        if not df_ag.empty:
            # Padronização de Data e Hora
            df_ag['DATA_DT'] = pd.to_datetime(df_ag['DATA'], errors='coerce', dayfirst=True).dt.date
            col_h = "HORARIO" if "HORARIO" in df_ag.columns else "HORA"
            
            # Cruzamento de dados para pegar o Endereço (A1_END)
            # Assume-se que 'CLIENTE' no Agendamento casa com 'A1_NOME' ou similar na Para_Agendar
            # Ajuste 'A1_NOME' se o nome da coluna de cliente na Para_Agendar for diferente
            if "A1_END" in df_cli.columns:
                dict_enderecos = pd.Series(df_cli.A1_END.values, index=df_cli.A1_NOME).to_dict()
                df_ag['ENDERECO_FMT'] = df_ag['CLIENTE'].map(dict_enderecos).fillna("Endereço não encontrado")
            else:
                df_ag['ENDERECO_FMT'] = "Coluna A1_END não localizada"

            df_ag = df_ag.sort_values(by=["DATA_DT", col_h])
            
        return df_ag
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return pd.DataFrame()

# --- 3. INTERFACE ---
st.title("📅 Calendário de Visitas")

# Controle de Meses (Simplificado para o exemplo)
if 'mes_ref' not in st.session_state:
    st.session_state.mes_ref = date.today().replace(day=1)

df_ag = carregar_dados_completos()

# Renderização do Calendário (Grid)
cal = calendar.monthcalendar(st.session_state.mes_ref.year, st.session_state.mes_ref.month)
cols_h = st.columns(7)
dias_semana = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
for i, d in enumerate(dias_semana):
    cols_h[i].markdown(f"<p style='text-align:center;font-weight:bold;'>{d}</p>", unsafe_allow_html=True)

for semana in cal:
    cols = st.columns(7)
    for i, dia in enumerate(semana):
        if dia == 0: continue
        with cols[i]:
            data_atual = date(st.session_state.mes_ref.year, st.session_state.mes_ref.month, dia)
            st.markdown(f"**{dia}**")
            
            if not df_ag.empty:
                visitas_dia = df_ag[df_ag['DATA_DT'] == data_atual]
                for idx, v in visitas_dia.iterrows():
                    h = v.get('HORARIO', v.get('HORA', '--:--'))
                    cli = v.get('CLIENTE', 'N/A')
                    
                    with st.expander(f"📍 {h} | {cli[:10]}"):
                        st.write(f"**👤 Cliente:** {cli}")
                        st.write(f"**🏠 Endereço:** {v.get('ENDERECO_FMT', 'N/A')}")
                        # CORREÇÃO: Pegando da coluna 'CONTATO' (Coluna I da planilha)
                        st.write(f"**📞 Contato:** {v.get('CONTATO', 'N/A')}")
                        st.write(f"**💰 Valor:** R$ {v.get('VALOR TOTAL', '0,00')}")
                        st.caption(f"📝 {v.get('FINALIDADE', '')}")

st.divider()
if st.button("⬅️ Voltar"):
    st.switch_page("main.py")
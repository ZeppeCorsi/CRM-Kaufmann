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

def carregar_dados_calendario():
    try:
        sh = conectar_google_sheets()
        
        # 1. Carrega aba Agendamentos (Onde está o Contato gravado)
        ws_ag = sh.worksheet("Agendamentos")
        df_ag = pd.DataFrame(ws_ag.get_all_records())
        df_ag.columns = [str(c).strip().upper() for c in df_ag.columns]
        
        # 2. Carrega aba Para_Agendar (Onde está o Endereço A1_END)
        ws_para = sh.worksheet("Para_Agendar")
        df_para = pd.DataFrame(ws_para.get_all_records())
        df_para.columns = [str(c).strip().upper() for c in df_para.columns]
        
        if not df_ag.empty:
            # Tratamento de Data e Hora
            df_ag['DATA_DT'] = pd.to_datetime(df_ag['DATA'], errors='coerce', dayfirst=True).dt.date
            col_h = "HORARIO" if "HORARIO" in df_ag.columns else "HORA"
            
            # --- LÓGICA DE PROCV (MERGE) ---
            # Relacionamos as duas abas pela coluna CLIENTE que é comum a ambas
            if "CLIENTE" in df_para.columns and "A1_END" in df_para.columns:
                # Pegamos apenas as colunas necessárias da aba Para_Agendar para não sujar o DF
                df_enderecos = df_para[['CLIENTE', 'A1_END']].drop_duplicates(subset=['CLIENTE'])
                # Fazemos o merge (Left Join)
                df_ag = pd.merge(df_ag, df_enderecos, on='CLIENTE', how='left')
            
            df_ag = df_ag.sort_values(by=["DATA_DT", col_h])
            
        return df_ag
    except Exception as e:
        st.error(f"Erro ao processar integração: {e}")
        return pd.DataFrame()

# --- 3. INTERFACE ---
st.title("📅 Calendário de Visitas")

if 'mes_ref' not in st.session_state:
    st.session_state.mes_ref = date.today().replace(day=1)

# Navegação de meses
col_nav1, col_nav2, col_nav3 = st.columns([1, 2, 1])
if col_nav1.button("⬅️ Anterior"):
    st.session_state.mes_ref = (st.session_state.mes_ref - timedelta(days=1)).replace(day=1)
    st.rerun()

meses_pt = {1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO", 4: "ABRIL", 5: "MAIO", 6: "JUNHO",
            7: "JULHO", 8: "AGOSTO", 9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO"}
col_nav2.markdown(f"<h3 style='text-align: center;'>{meses_pt[st.session_state.mes_ref.month]} {st.session_state.mes_ref.year}</h3>", unsafe_allow_html=True)

if col_nav3.button("Próximo ➡️"):
    st.session_state.mes_ref = (st.session_state.mes_ref + timedelta(days=32)).replace(day=1)
    st.rerun()

df_ag = carregar_dados_calendario()

# Grid do Calendário
cal = calendar.monthcalendar(st.session_state.mes_ref.year, st.session_state.mes_ref.month)
dias_semana = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
cols_h = st.columns(7)
for i, d in enumerate(dias_semana):
    cols_h[i].markdown(f"<p style='text-align:center;font-weight:bold;color:#555;'>{d}</p>", unsafe_allow_html=True)

for semana in cal:
    cols = st.columns(7)
    for i, dia in enumerate(semana):
        if dia == 0: continue
        with cols[i]:
            data_atual = date(st.session_state.mes_ref.year, st.session_state.mes_ref.month, dia)
            cor_dia = "blue" if data_atual == date.today() else "black"
            st.markdown(f"<p style='text-align:left; font-weight:bold; color:{cor_dia};'>{dia}</p>", unsafe_allow_html=True)
            
            if not df_ag.empty:
                visitas_dia = df_ag[df_ag['DATA_DT'] == data_atual]
                
                for idx, v in visitas_dia.iterrows():
                    realizada = str(v.get('REALIZADA', '')).upper() == "SIM"
                    icone = "✅" if realizada else "📍"
                    h = v.get('HORARIO', v.get('HORA', '--:--'))
                    cli = v.get('CLIENTE', 'N/A')
                    
                    label = f"{icone} {h} | {cli[:10]}"
                    with st.expander(label):
                        st.write(f"**👤 Cliente:** {cli}")
                        
                        # Endereço vindo da aba Para_Agendar via merge
                        end = v.get('A1_END', 'Endereço não cadastrado')
                        st.write(f"**🏠 Endereço:** {end}")
                        
                        # Contato vindo diretamente da aba Agendamentos (Coluna I)
                        contato_gravado = v.get('CONTATO', 'N/A')
                        st.write(f"**📞 Contato:** {contato_gravado}")
                        
                        st.write(f"**💰 Valor:** R$ {v.get('VALOR TOTAL', '0,00')}")
                        st.caption(f"📝 Finalidade: {v.get('FINALIDADE', 'N/A')}")

st.divider()
if st.button("⬅️ Voltar"):
    st.switch_page("main.py")
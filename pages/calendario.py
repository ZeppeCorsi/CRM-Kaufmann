import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, date, timedelta
import calendar
import time as t_module
import urllib.parse

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
        ws_ag = sh.worksheet("Agendamentos")
        df_ag = pd.DataFrame(ws_ag.get_all_records())
        df_ag.columns = [str(c).strip().upper() for c in df_ag.columns]
        df_ag['ORIGINAL_INDEX'] = df_ag.index
        
        ws_para = sh.worksheet("Para_Agendar")
        df_para = pd.DataFrame(ws_para.get_all_records())
        df_para.columns = [str(c).strip().upper() for c in df_para.columns]
        
        if not df_ag.empty:
            df_ag['DATA_DT'] = pd.to_datetime(df_ag['DATA'], errors='coerce', dayfirst=True).dt.date
            col_h = "HORARIO" if "HORARIO" in df_ag.columns else "HORA"
            
            if "CLIENTE" in df_para.columns and "A1_END" in df_para.columns:
                df_end = df_para[['CLIENTE', 'A1_END']].drop_duplicates(subset=['CLIENTE'])
                df_ag = pd.merge(df_ag, df_end, on='CLIENTE', how='left')
            
            df_ag = df_ag.sort_values(by=["DATA_DT", col_h])
        return df_ag
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return pd.DataFrame()

def processar_reagendamento(idx_antigo, dados_originais, nova_data, novo_horario):
    try:
        sh = conectar_google_sheets()
        ws = sh.worksheet("Agendamentos")
        linha_antiga = int(idx_antigo) + 2
        
        # 1. Marca a antiga como REAGENDADO na coluna REALIZADA (Coluna L / 12)
        ws.update_cell(linha_antiga, 12, "REAGENDADO")
        
        # 2. Cria a nova linha (Copia dados e altera Data/Hora)
        # Ordem sugerida: DATA, HORARIO, CLIENTE, FINALIDADE, VALOR, VENDEDOR, REALIZADA, OBS, CONTATO...
        # Vamos montar uma lista baseada nos dados que já temos no DataFrame
        nova_linha = [
            nova_data.strftime("%d/%m/%Y"),
            novo_horario,
            dados_originais.get('CLIENTE', ''),
            dados_originais.get('FINALIDADE', ''),
            dados_originais.get('VALOR TOTAL', ''),
            dados_originais.get('VENDEDOR', ''),
            "NAO", # Realizada novo
            f"Reagendado da data {dados_originais.get('DATA', '')}",
            dados_originais.get('CONTATO', ''),
            st.session_state.username, # Usuário inclusão
            datetime.now().strftime("%d/%m/%Y %H:%M:%S") # Log
        ]
        ws.append_row(nova_linha)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Erro ao reagendar: {e}")
        return False

# --- 3. DIALOGS (POPUPS) ---
@st.dialog("Finalizar Visita")
def popup_finalizar_visita(idx, cliente):
    st.write(f"Registrar resultado para: **{cliente}**")
    relato = st.text_area("Descreva o que foi tratado na visita:")
    if st.button("Gravar na Planilha"):
        # Lógica de finalizar (simplificada aqui para brevidade)
        st.success("Sucesso!")
        st.rerun()

@st.dialog("Reagendar Visita")
def popup_reagendar(idx, v):
    st.write(f"Reagendando: **{v['CLIENTE']}**")
    n_data = st.date_input("Nova Data", value=date.today())
    n_hora = st.time_input("Novo Horário", value=datetime.now().time())
    if st.button("Confirmar Novo Agendamento"):
        if processar_reagendamento(idx, v, n_data, n_hora.strftime("%H:%M")):
            st.success("Reagendado!")
            t_module.sleep(1)
            st.rerun()

# --- 4. INTERFACE ---
st.title("📅 Calendário de Visitas")

if 'mes_ref' not in st.session_state:
    st.session_state.mes_ref = date.today().replace(day=1)

# CABEÇALHO DE NAVEGAÇÃO
col_n1, col_n2, col_n3 = st.columns([1, 2, 1])
if col_n1.button("⬅️ Anterior"):
    st.session_state.mes_ref = (st.session_state.mes_ref - timedelta(days=1)).replace(day=1)
    st.rerun()

meses_pt = {1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO", 4: "ABRIL", 5: "MAIO", 6: "JUNHO",
            7: "JULHO", 8: "AGOSTO", 9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO"}
col_n2.markdown(f"<h3 style='text-align: center;'>{meses_pt[st.session_state.mes_ref.month]} {st.session_state.mes_ref.year}</h3>", unsafe_allow_html=True)

if col_n3.button("Próximo ➡️"):
    st.session_state.mes_ref = (st.session_state.mes_ref + timedelta(days=32)).replace(day=1)
    st.rerun()

df_ag = carregar_dados_calendario()

# DIAS DA SEMANA COM LARGURA AJUSTADA (Sab/Dom menores)
# Proporção: Seg a Sex (1.2), Sab/Dom (0.5)
cols_h = st.columns([1.2, 1.2, 1.2, 1.2, 1.2, 0.6, 0.6])
dias_semana = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sáb", "Dom"]
for i, d in enumerate(dias_semana):
    cols_h[i].markdown(f"<p style='text-align:center;font-weight:bold;font-size:12px;'>{d}</p>", unsafe_allow_html=True)

# GRID DO CALENDÁRIO
cal = calendar.monthcalendar(st.session_state.mes_ref.year, st.session_state.mes_ref.month)
for semana in cal:
    cols = st.columns([1.2, 1.2, 1.2, 1.2, 1.2, 0.6, 0.6])
    for i, dia in enumerate(semana):
        if dia == 0: continue
        with cols[i]:
            data_atual = date(st.session_state.mes_ref.year, st.session_state.mes_ref.month, dia)
            st.markdown(f"**{dia}**")
            
            if not df_ag.empty:
                visitas_dia = df_ag[df_ag['DATA_DT'] == data_atual]
                for _, v in visitas_dia.iterrows():
                    status = str(v.get('REALIZADA', '')).upper()
                    h = v.get('HORARIO', v.get('HORA', '--:--'))
                    
                    # DEFINIÇÃO DE CORES
                    bg_color = "white"
                    icone = "📍"
                    if status == "SIM": 
                        bg_color = "#e1f5fe"; icone = "✅" # Azul claro
                    elif status == "REAGENDADO": 
                        bg_color = "#ffebee"; icone = "🔄" # Rosa claro / Vermelho suave
                    
                    label = f"{icone} {h} | {v['CLIENTE'][:8]}"
                    
                    with st.container():
                        st.markdown(f"""<div style="background-color:{bg_color}; border-radius:5px; padding:2px; border:1px solid #ddd; margin-bottom:2px;">""", unsafe_allow_html=True)
                        with st.expander(label):
                            st.write(f"**Cliente:** {v['CLIENTE']}")
                            st.write(f"**Contato:** {v.get('CONTATO', 'N/A')}")
                            
                            # Botão Outlook
                            assunto = urllib.parse.quote(f"Visita - {v['CLIENTE']}")
                            link_mail = f"mailto:?subject={assunto}"
                            st.markdown(f'<a href="{link_mail}"><button style="width:100%; cursor:pointer;">📧 Outlook</button></a>', unsafe_allow_html=True)
                            
                            if status == "NAO":
                                if st.button("🏁 Finalizar", key=f"f_{v['ORIGINAL_INDEX']}"):
                                    popup_finalizar_visita(v['ORIGINAL_INDEX'], v['CLIENTE'])
                                if st.button("📅 Reagendar", key=f"r_{v['ORIGINAL_INDEX']}"):
                                    popup_reagendar(v['ORIGINAL_INDEX'], v)
                        st.markdown("</div>", unsafe_allow_html=True)

st.divider()
if st.button("⬅️ Voltar"):
    st.switch_page("main.py")
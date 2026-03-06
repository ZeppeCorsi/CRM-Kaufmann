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
        # Carrega Agendamentos
        ws_ag = sh.worksheet("Agendamentos")
        df_ag = pd.DataFrame(ws_ag.get_all_records())
        df_ag.columns = [str(c).strip().upper() for c in df_ag.columns]
        df_ag['ORIGINAL_INDEX'] = df_ag.index
        
        # Carrega Para_Agendar para buscar o Endereço (A1_END)
        ws_para = sh.worksheet("Para_Agendar")
        df_para = pd.DataFrame(ws_para.get_all_records())
        df_para.columns = [str(c).strip().upper() for c in df_para.columns]
        
        if not df_ag.empty:
            df_ag['DATA_DT'] = pd.to_datetime(df_ag['DATA'], errors='coerce', dayfirst=True).dt.date
            col_h = "HORARIO" if "HORARIO" in df_ag.columns else "HORA"
            
            # Cruzamento de abas pelo CLIENTE
            if "CLIENTE" in df_para.columns and "A1_END" in df_para.columns:
                df_end = df_para[['CLIENTE', 'A1_END']].drop_duplicates(subset=['CLIENTE'])
                df_ag = pd.merge(df_ag, df_end, on='CLIENTE', how='left')
            
            df_ag = df_ag.sort_values(by=["DATA_DT", col_h])
        return df_ag
    except Exception as e:
        st.error(f"Erro ao carregar: {e}")
        return pd.DataFrame()

# --- 3. POPUPS (DIALOGS) ---

@st.dialog("Finalizar Visita")
def popup_finalizar_visita(idx, cliente):
    st.subheader(f"🏁 Finalizar Visita: {cliente}")
    relato = st.text_area("Notas da visita:", placeholder="O que foi acordado?")
    if st.button("Gravar Resultado"):
        try:
            sh = conectar_google_sheets()
            ws = sh.worksheet("Agendamentos")
            linha = int(idx) + 2
            ws.update_cell(linha, 12, "SIM") # Coluna L: REALIZADA
            ws.update_cell(linha, 8, f"RELATO: {relato}") # Coluna H: OBS
            st.success("Visita Finalizada!")
            st.cache_data.clear()
            t_module.sleep(1)
            st.rerun()
        except Exception as e:
            st.error(f"Erro: {e}")

@st.dialog("Reagendar Visita")
def popup_reagendar(idx, v):
    st.subheader(f"📅 Reagendar: {v['CLIENTE']}")
    n_data = st.date_input("Nova Data", value=date.today())
    n_hora = st.time_input("Novo Horário")
    if st.button("Confirmar Reagendamento"):
        try:
            sh = conectar_google_sheets()
            ws = sh.worksheet("Agendamentos")
            linha_antiga = int(idx) + 2
            ws.update_cell(linha_antiga, 12, "REAGENDADO")
            
            nova_linha = [
                n_data.strftime("%d/%m/%Y"),
                n_hora.strftime("%H:%M"),
                v.get('CLIENTE', ''),
                v.get('FINALIDADE', ''),
                v.get('VALOR TOTAL', ''),
                v.get('VENDEDOR', ''),
                "NAO",
                f"Antiga visita de {v.get('DATA', '')}",
                v.get('CONTATO', ''),
                st.session_state.username,
                datetime.now().strftime("%d/%m/%Y %H:%M")
            ]
            ws.append_row(nova_linha)
            st.success("Reagendado!")
            st.cache_data.clear()
            t_module.sleep(1)
            st.rerun()
        except Exception as e:
            st.error(f"Erro: {e}")

# --- 4. INTERFACE ---

st.title("📅 Calendário de Visitas")

if 'mes_ref' not in st.session_state:
    st.session_state.mes_ref = date.today().replace(day=1)

# Navegação
c1, c2, c3 = st.columns([1, 2, 1])
if c1.button("⬅️ Anterior"):
    st.session_state.mes_ref = (st.session_state.mes_ref - timedelta(days=1)).replace(day=1)
    st.rerun()

meses_pt = {1:"JANEIRO", 2:"FEVEREIRO", 3:"MARÇO", 4:"ABRIL", 5:"MAIO", 6:"JUNHO",
            7:"JULHO", 8:"AGOSTO", 9:"SETEMBRO", 10:"OUTUBRO", 11:"NOVEMBRO", 12:"DEZEMBRO"}
c2.markdown(f"<h3 style='text-align: center;'>{meses_pt[st.session_state.mes_ref.month]} {st.session_state.mes_ref.year}</h3>", unsafe_allow_html=True)

if c3.button("Próximo ➡️"):
    st.session_state.mes_ref = (st.session_state.mes_ref + timedelta(days=32)).replace(day=1)
    st.rerun()

df_ag = carregar_dados_calendario()

# Dias da Semana
cols_h = st.columns([1.2, 1.2, 1.2, 1.2, 1.2, 0.7, 0.7])
dias = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]
for i, d in enumerate(dias):
    cols_h[i].markdown(f"<p style='text-align:center; font-weight:bold;'>{d}</p>", unsafe_allow_html=True)

# Grid
cal = calendar.monthcalendar(st.session_state.mes_ref.year, st.session_state.mes_ref.month)
for semana in cal:
    cols = st.columns([1.2, 1.2, 1.2, 1.2, 1.2, 0.7, 0.7])
    for i, dia in enumerate(semana):
        if dia == 0: continue
        with cols[i]:
            data_atual = date(st.session_state.mes_ref.year, st.session_state.mes_ref.month, dia)
            st.markdown(f"**{dia}**")
            
            if not df_ag.empty:
                visitas = df_ag[df_ag['DATA_DT'] == data_atual]
                for _, v in visitas.iterrows():
                    status = str(v.get('REALIZADA', '')).upper()
                    cli = v.get('CLIENTE', 'N/A')
                    hora = v.get('HORARIO', v.get('HORA', '--:--'))
                    end = v.get('A1_END', 'Endereço não localizado') #
                    cont = v.get('CONTATO', 'N/A') #
                    
                    bg = "white"; icone = "📍"
                    if status == "SIM": bg = "#E8F5E9"; icone = "✅"
                    elif status == "REAGENDADO": bg = "#FFEBEE"; icone = "🔄"
                    
                    st.markdown(f'<div style="background-color:{bg}; padding:5px; border-radius:5px; border:1px solid #ddd; margin-bottom:5px;">', unsafe_allow_html=True)
                    with st.expander(f"{icone} {hora} | {cli[:8]}"):
                        st.write(f"**Cliente:** {cli}")
                        st.write(f"**Endereço:** {end}")
                        st.write(f"**Contato:** {cont}")
                        
                        # Botão Outlook Azul
                        assunto = urllib.parse.quote(f"Visita: {cli}")
                        corpo = urllib.parse.quote(f"Cliente: {cli}\nEndereço: {end}\nContato: {cont}")
                        st.markdown(f"""<a href="mailto:?subject={assunto}&body={corpo}" target="_blank">
                                        <button style="width:100%; background-color:#0078d4; color:white; border:none; padding:8px; border-radius:5px; cursor:pointer; margin-bottom:10px;">📧 Enviar Outlook</button>
                                    </a>""", unsafe_allow_html=True)
                        
                        # MOSTRAR SEMPRE SE NÃO ESTIVER FINALIZADA
                        if status == "NAO" or status == "":
                            st.write("---")
                            # BOTÃO FINALIZAR
                            if st.button("🏁 Finalizar", key=f"f_{v['ORIGINAL_INDEX']}"):
                                popup_finalizar_visita(v['ORIGINAL_INDEX'], cli)
                            # BOTÃO REAGENDAR
                            if st.button("📅 Reagendar", key=f"r_{v['ORIGINAL_INDEX']}"):
                                popup_reagendar(v['ORIGINAL_INDEX'], v)
                    st.markdown('</div>', unsafe_allow_html=True)

st.divider()
if st.button("⬅️ Sair"): st.switch_page("main.py")
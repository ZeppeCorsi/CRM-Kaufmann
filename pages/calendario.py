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
            
            # MERGE PARA O ENDEREÇO (A1_END) - MANTIDO!
            if "CLIENTE" in df_para.columns and "A1_END" in df_para.columns:
                df_end = df_para[['CLIENTE', 'A1_END']].drop_duplicates(subset=['CLIENTE'])
                df_ag = pd.merge(df_ag, df_end, on='CLIENTE', how='left')
            
            df_ag = df_ag.sort_values(by=["DATA_DT", col_h])
        return df_ag
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return pd.DataFrame()

def processar_reagendamento(idx_antigo, v, nova_data, novo_horario):
    try:
        sh = conectar_google_sheets()
        ws = sh.worksheet("Agendamentos")
        linha_antiga = int(idx_antigo) + 2
        
        # 1. Marca a antiga como REAGENDADO
        ws.update_cell(linha_antiga, 12, "REAGENDADO") # Coluna L
        
        # 2. Cria a nova linha com a nova data e hora
        nova_linha = [
            nova_data.strftime("%d/%m/%Y"),
            novo_horario,
            v.get('CLIENTE', ''),
            v.get('FINALIDADE', ''),
            v.get('VALOR TOTAL', ''),
            v.get('VENDEDOR', ''),
            "NAO", 
            f"Reagendado de {v.get('DATA', '')}",
            v.get('CONTATO', ''),
            st.session_state.username,
            datetime.now().strftime("%d/%m/%Y %H:%M:%S")
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
        sh = conectar_google_sheets()
        ws = sh.worksheet("Agendamentos")
        linha = int(idx) + 2
        ws.update_cell(linha, 12, "SIM")
        ws.update_cell(linha, 8, relato)
        st.success("Visita Finalizada!")
        t_module.sleep(1)
        st.rerun()

@st.dialog("Reagendar Visita")
def popup_reagendar(idx, v):
    st.write(f"Reagendando: **{v['CLIENTE']}**")
    n_data = st.date_input("Nova Data", value=date.today())
    n_hora = st.time_input("Novo Horário")
    if st.button("Confirmar Reagendamento"):
        if processar_reagendamento(idx, v, n_data, n_hora.strftime("%H:%M")):
            st.success("Reagendado com sucesso!")
            t_module.sleep(1)
            st.rerun()

# --- 4. INTERFACE ---
st.title("📅 Calendário de Visitas")

if 'mes_ref' not in st.session_state:
    st.session_state.mes_ref = date.today().replace(day=1)

# Navegação
col_n1, col_n2, col_n3 = st.columns([1, 2, 1])
if col_n1.button("⬅️ Anterior"):
    st.session_state.mes_ref = (st.session_state.mes_ref - timedelta(days=1)).replace(day=1)
    st.rerun()

meses_pt = {i+1: m for i, m in enumerate(["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"])}
col_n2.markdown(f"<h3 style='text-align: center;'>{meses_pt[st.session_state.mes_ref.month]} {st.session_state.mes_ref.year}</h3>", unsafe_allow_html=True)

if col_n3.button("Próximo ➡️"):
    st.session_state.mes_ref = (st.session_state.mes_ref + timedelta(days=32)).replace(day=1)
    st.rerun()

df_ag = carregar_dados_calendario()

# Dias da semana com colunas ajustadas
cols_h = st.columns([1.2, 1.2, 1.2, 1.2, 1.2, 0.7, 0.7])
dias = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]
for i, d in enumerate(dias):
    cols_h[i].markdown(f"<p style='text-align:center;font-weight:bold;'>{d}</p>", unsafe_allow_html=True)

cal = calendar.monthcalendar(st.session_state.mes_ref.year, st.session_state.mes_ref.month)
for semana in cal:
    cols = st.columns([1.2, 1.2, 1.2, 1.2, 1.2, 0.7, 0.7])
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
                    cli = v.get('CLIENTE', 'N/A')
                    end = v.get('A1_END', 'Não localizado')
                    contato = v.get('CONTATO', 'N/A')
                    
                    # Definição de Cores
                    bg_color = "#ffffff" # Pendente
                    icone = "📍"
                    if status == "SIM": bg_color = "#e8f5e9"; icone = "✅" # Verde claro
                    elif status == "REAGENDADO": bg_color = "#ffebee"; icone = "🔄" # Rosa/Vermelho claro
                    
                    # Estilo do Card
                    st.markdown(f"""<div style="background-color:{bg_color}; border-radius:5px; padding:3px; border:1px solid #ddd; margin-bottom:5px;">""", unsafe_allow_html=True)
                    with st.expander(f"{icone} {h} | {cli[:8]}"):
                        st.write(f"**👤 Cliente:** {cli}")
                        st.write(f"**🏠 Endereço:** {end}")
                        st.write(f"**📞 Contato:** {contato}")
                        
                        # Botão Outlook Estilizado
                        corpo_email = f"Cliente: {cli}\nEndereço: {end}\nContato: {contato}"
                        link_mail = f"mailto:?subject=Visita {cli}&body={urllib.parse.quote(corpo_email)}"
                        st.markdown(f'<a href="{link_mail}" target="_blank"><button style="width:100%; background-color:#0078d4; color:white; border:none; padding:5px; border-radius:3px; cursor:pointer;">📧 Outlook</button></a>', unsafe_allow_html=True)
                        
                        if status == "NAO":
                            st.write("---")
                            if st.button("🏁 Finalizar", key=f"f_{v['ORIGINAL_INDEX']}"):
                                popup_finalizar_visita(v['ORIGINAL_INDEX'], cli)
                            if st.button("📅 Reagendar", key=f"r_{v['ORIGINAL_INDEX']}"):
                                popup_reagendar(v['ORIGINAL_INDEX'], v)
                    st.markdown("</div>", unsafe_allow_html=True)

st.divider()
if st.button("⬅️ Voltar"): st.switch_page("main.py")
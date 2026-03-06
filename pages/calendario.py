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
        # Carrega Agendamentos (Onde está o Contato e Status)
        ws_ag = sh.worksheet("Agendamentos")
        df_ag = pd.DataFrame(ws_ag.get_all_records())
        df_ag.columns = [str(c).strip().upper() for c in df_ag.columns]
        df_ag['ORIGINAL_INDEX'] = df_ag.index
        
        # Carrega Para_Agendar (Onde está o Endereço A1_END)
        ws_para = sh.worksheet("Para_Agendar")
        df_para = pd.DataFrame(ws_para.get_all_records())
        df_para.columns = [str(c).strip().upper() for c in df_para.columns]
        
        if not df_ag.empty:
            df_ag['DATA_DT'] = pd.to_datetime(df_ag['DATA'], errors='coerce', dayfirst=True).dt.date
            col_h = "HORARIO" if "HORARIO" in df_ag.columns else "HORA"
            
            # Cruzamento para buscar o Endereço (A1_END)
            if "CLIENTE" in df_para.columns and "A1_END" in df_para.columns:
                df_end = df_para[['CLIENTE', 'A1_END']].drop_duplicates(subset=['CLIENTE'])
                df_ag = pd.merge(df_ag, df_end, on='CLIENTE', how='left')
            
            df_ag = df_ag.sort_values(by=["DATA_DT", col_h])
        return df_ag
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return pd.DataFrame()

# --- 3. POPUPS (DIALOGS) ---

@st.dialog("Finalizar Visita")
def popup_finalizar_visita(idx, cliente):
    st.subheader(f"🏁 Finalizar: {cliente}")
    relato = st.text_area("Descreva o resultado da visita:", height=100)
    if st.button("Confirmar Finalização"):
        try:
            sh = conectar_google_sheets()
            ws = sh.worksheet("Agendamentos")
            linha = int(idx) + 2
            ws.update_cell(linha, 12, "SIM") # Coluna L: REALIZADA
            # Salva o relato na coluna H (ENDERECO/OBS)
            ws.update_cell(linha, 8, f"RESULTADO: {relato}")
            st.success("Visita finalizada!")
            st.cache_data.clear()
            t_module.sleep(1)
            st.rerun()
        except Exception as e:
            st.error(f"Erro: {e}")

@st.dialog("Reagendar Visita")
def popup_reagendar(idx, v):
    st.subheader(f"📅 Reagendar: {v['CLIENTE']}")
    nova_data = st.date_input("Selecione a nova data", value=date.today())
    nova_hora = st.time_input("Selecione o novo horário")
    if st.button("Salvar Reagendamento"):
        try:
            sh = conectar_google_sheets()
            ws = sh.worksheet("Agendamentos")
            linha_antiga = int(idx) + 2
            ws.update_cell(linha_antiga, 12, "REAGENDADO") # Marca antiga como REAGENDADO
            
            # Cria nova linha com os dados
            nova_linha = [
                nova_data.strftime("%d/%m/%Y"),
                nova_hora.strftime("%H:%M"),
                v.get('CLIENTE', ''),
                v.get('FINALIDADE', ''),
                v.get('VALOR TOTAL', ''),
                v.get('VENDEDOR', ''),
                "NAO",
                f"Reagendado de {v.get('DATA','')}",
                v.get('CONTATO', ''),
                st.session_state.username,
                datetime.now().strftime("%d/%m/%Y %H:%M")
            ]
            ws.append_row(nova_linha)
            st.success("Nova visita agendada!")
            st.cache_data.clear()
            t_module.sleep(1)
            st.rerun()
        except Exception as e:
            st.error(f"Erro: {e}")

# --- 4. INTERFACE DO CALENDÁRIO ---

st.title("📅 Calendário de Visitas")

if 'mes_ref' not in st.session_state:
    st.session_state.mes_ref = date.today().replace(day=1)

# Navegação de Meses
col_n1, col_n2, col_n3 = st.columns([1, 2, 1])
if col_n1.button("⬅️ Mês Anterior"):
    st.session_state.mes_ref = (st.session_state.mes_ref - timedelta(days=1)).replace(day=1)
    st.rerun()

meses_dict = {1:"JANEIRO", 2:"FEVEREIRO", 3:"MARÇO", 4:"ABRIL", 5:"MAIO", 6:"JUNHO",
              7:"JULHO", 8:"AGOSTO", 9:"SETEMBRO", 10:"OUTUBRO", 11:"NOVEMBRO", 12:"DEZEMBRO"}
col_n2.markdown(f"<h3 style='text-align: center;'>{meses_dict[st.session_state.mes_ref.month]} {st.session_state.mes_ref.year}</h3>", unsafe_allow_html=True)

if col_n3.button("Próximo Mês ➡️"):
    st.session_state.mes_ref = (st.session_state.mes_ref + timedelta(days=32)).replace(day=1)
    st.rerun()

df_ag = carregar_dados_calendario()

# Cabeçalho dos dias com Sábado e Domingo menores
cols_h = st.columns([1.2, 1.2, 1.2, 1.2, 1.2, 0.7, 0.7])
dias_nomes = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sáb", "Dom"]
for i, nome in enumerate(dias_nomes):
    cols_h[i].markdown(f"<p style='text-align:center; font-weight:bold;'>{nome}</p>", unsafe_allow_html=True)

# Geração das semanas
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
                    hora = v.get('HORARIO', v.get('HORA', '--:--'))
                    cliente = v.get('CLIENTE', 'N/A')
                    endereco = v.get('A1_END', 'Não localizado')
                    contato = v.get('CONTATO', 'N/A')
                    
                    # Definir cores e ícones
                    bg_color = "white"; icone = "📍"
                    if status == "SIM": 
                        bg_color = "#E8F5E9"; icone = "✅"
                    elif status == "REAGENDADO": 
                        bg_color = "#FFEBEE"; icone = "🔄"
                    
                    st.markdown(f'<div style="background-color:{bg_color}; padding:5px; border-radius:5px; border:1px solid #ddd; margin-bottom:5px;">', unsafe_allow_html=True)
                    with st.expander(f"{icone} {hora} | {cliente[:10]}"):
                        st.write(f"**Cliente:** {cliente}")
                        st.write(f"**Endereço:** {endereco}")
                        st.write(f"**Contato:** {contato}")
                        
                        # Botão Outlook Azul
                        assunto = urllib.parse.quote(f"Visita: {cliente}")
                        corpo = urllib.parse.quote(f"Cliente: {cliente}\nEndereço: {endereco}\nContato: {contato}")
                        st.markdown(f"""
                            <a href="mailto:?subject={assunto}&body={corpo}" target="_blank">
                                <button style="width:100%; background-color:#0078d4; color:white; border:none; padding:8px; border-radius:5px; cursor:pointer; margin-bottom:10px;">
                                    📧 Enviar Outlook
                                </button>
                            </a>""", unsafe_allow_html=True)
                        
                        if status == "NAO":
                            # Botão Finalizar
                            if st.button("🏁 Finalizar", key=f"fin_{v['ORIGINAL_INDEX']}"):
                                popup_finalizar_visita(v['ORIGINAL_INDEX'], cliente)
                            
                            # Botão Reagendar
                            if st.button("📅 Reagendar", key=f"re_{v['ORIGINAL_INDEX']}"):
                                popup_reagendar(v['ORIGINAL_INDEX'], v)
                    st.markdown('</div>', unsafe_allow_html=True)

st.divider()
if st.button("⬅️ Voltar"): st.switch_page("main.py")
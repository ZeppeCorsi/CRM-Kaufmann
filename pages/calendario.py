import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, date, timedelta
import calendar
import time as t_module
import urllib.parse  # Para formatar o link do e-mail

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

def atualizar_visita_gs(indice_original, novo_orc_resp, data_follow, novos_detalhes):
    try:
        sh = conectar_google_sheets()
        worksheet = sh.worksheet("Agendamentos")
        linha = int(indice_original) + 2
        worksheet.update_cell(linha, 12, "SIM") # Coluna L: REALIZADA
        
        obs_atual = worksheet.cell(linha, 8).value or ""
        nova_obs = f"{obs_atual} | RESULTADO: {novos_detalhes}".strip(" | ")
        worksheet.update_cell(linha, 8, nova_obs) 
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Erro ao finalizar: {e}")
        return False

# --- 3. POPUP DE FINALIZAÇÃO ---
@st.dialog("Finalizar Visita")
def popup_finalizar_visita(idx, cliente):
    st.write(f"Registrar resultado para: **{cliente}**")
    novo_orc = st.radio("Gerou um NOVO Orçamento?", ["SIM", "NAO"], horizontal=True)
    follow = st.date_input("Próxima Data de Follow-up", value=date.today() + timedelta(days=7), format="DD/MM/YYYY")
    relato = st.text_area("Descreva o que foi tratado na visita:")
    
    if st.button("Gravar na Planilha"):
        if atualizar_visita_gs(idx, novo_orc, follow, relato):
            st.balloons()
            st.success("✅ Visita finalizada com sucesso!")
            t_module.sleep(1)
            st.rerun()

# --- 4. INTERFACE ---
st.title("📅 Calendário de Visitas")

if 'mes_ref' not in st.session_state:
    st.session_state.mes_ref = date.today().replace(day=1)

# Navegação (omitida para brevidade, mas mantida no seu código original)
df_ag = carregar_dados_calendario()

cal = calendar.monthcalendar(st.session_state.mes_ref.year, st.session_state.mes_ref.month)
for semana in cal:
    cols = st.columns(7)
    for i, dia in enumerate(semana):
        if dia == 0: continue
        with cols[i]:
            data_atual = date(st.session_state.mes_ref.year, st.session_state.mes_ref.month, dia)
            st.markdown(f"**{dia}**")
            
            if not df_ag.empty:
                visitas_dia = df_ag[df_ag['DATA_DT'] == data_atual]
                for _, v in visitas_dia.iterrows():
                    realizada = str(v.get('REALIZADA', '')).upper() == "SIM"
                    h = v.get('HORARIO', v.get('HORA', '--:--'))
                    cli = v.get('CLIENTE', 'N/A')
                    contato = v.get('CONTATO', 'N/A')
                    endereco = v.get('A1_END', 'N/A')
                    
                    with st.expander(f"{'✅' if realizada else '📍'} {h} | {cli[:10]}"):
                        st.write(f"**👤 Cliente:** {cli}")
                        st.write(f"**🏠 Endereço:** {endereco}")
                        st.write(f"**📞 Contato:** {contato}")
                        
                        # --- NOVO: BOTÃO ENVIAR PARA OUTLOOK ---
                        corpo_email = (
                            f"Detalhes da Visita Agendada:\n\n"
                            f"Cliente: {cli}\n"
                            f"Data: {data_atual.strftime('%d/%m/%Y')}\n"
                            f"Horário: {h}\n"
                            f"Contato: {contato}\n"
                            f"Endereço: {endereco}\n"
                            f"Finalidade: {v.get('FINALIDADE', 'N/A')}\n"
                        )
                        
                        # Formata o link mailto
                        assunto = urllib.parse.quote(f"Visita Comercial - {cli}")
                        corpo = urllib.parse.quote(corpo_email)
                        # Você pode colocar um e-mail padrão em mailto:email@exemplo.com
                        link_outlook = f"mailto:?subject={assunto}&body={corpo}"
                        
                        st.markdown(f"""
                            <a href="{link_outlook}" target="_blank" style="text-decoration: none;">
                                <button style="width: 100%; background-color: #0078d4; color: white; border: none; padding: 8px; border-radius: 5px; cursor: pointer; margin-bottom: 5px;">
                                    📧 Enviar p/ Outlook
                                </button>
                            </a>
                        """, unsafe_allow_html=True)

                        if not realizada:
                            if st.button("🏁 Finalizar", key=f"fin_{v['ORIGINAL_INDEX']}"):
                                popup_finalizar_visita(v['ORIGINAL_INDEX'], cli)
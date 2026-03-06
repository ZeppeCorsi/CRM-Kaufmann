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
        
        # 1. Carrega aba Agendamentos
        ws_ag = sh.worksheet("Agendamentos")
        data_ag = ws_ag.get_all_records()
        df_ag = pd.DataFrame(data_ag)
        df_ag.columns = [str(c).strip().upper() for c in df_ag.columns]
        
        # Guardamos o índice original para poder atualizar a linha correta depois
        df_ag['ORIGINAL_INDEX'] = df_ag.index
        
        # 2. Carrega aba Para_Agendar para o endereço
        ws_para = sh.worksheet("Para_Agendar")
        df_para = pd.DataFrame(ws_para.get_all_records())
        df_para.columns = [str(c).strip().upper() for c in df_para.columns]
        
        if not df_ag.empty:
            df_ag['DATA_DT'] = pd.to_datetime(df_ag['DATA'], errors='coerce', dayfirst=True).dt.date
            col_h = "HORARIO" if "HORARIO" in df_ag.columns else "HORA"
            
            # MERGE para buscar endereço A1_END
            if "CLIENTE" in df_para.columns and "A1_END" in df_para.columns:
                df_end = df_para[['CLIENTE', 'A1_END']].drop_duplicates(subset=['CLIENTE'])
                df_ag = pd.merge(df_ag, df_end, on='CLIENTE', how='left')
            
            df_ag = df_ag.sort_values(by=["DATA_DT", col_h])
            
        return df_ag
    except Exception as e:
        st.error(f"Erro ao processar dados: {e}")
        return pd.DataFrame()

def atualizar_visita_gs(indice_original, novo_orc_resp, data_follow, novos_detalhes):
    try:
        sh = conectar_google_sheets()
        worksheet = sh.worksheet("Agendamentos")
        # +2 porque gspread começa em 1 e tem cabeçalho
        linha = int(indice_original) + 2
        
        # Coluna L: REALIZADA (Baseado na imagem da planilha)
        worksheet.update_cell(linha, 12, "SIM") 
        
        # Coluna H: ENDERECO / OBS (Concatenando resultado)
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
                
                for _, v in visitas_dia.iterrows():
                    realizada = str(v.get('REALIZADA', '')).upper() == "SIM"
                    icone = "✅" if realizada else "📍"
                    h = v.get('HORARIO', v.get('HORA', '--:--'))
                    cli = v.get('CLIENTE', 'N/A')
                    
                    label = f"{icone} {h} | {cli[:10]}"
                    with st.expander(label):
                        st.write(f"**👤 Cliente:** {cli}")
                        st.write(f"**🏠 Endereço:** {v.get('A1_END', 'Endereço não localizado')}")
                        st.write(f"**📞 Contato:** {v.get('CONTATO', 'N/A')}")
                        st.write(f"**💰 Valor:** R$ {v.get('VALOR TOTAL', '0,00')}")
                        
                        if not realizada:
                            if st.button("Finalizar", key=f"fin_{v['ORIGINAL_INDEX']}"):
                                popup_finalizar_visita(v['ORIGINAL_INDEX'], cli)

st.divider()
if st.button("⬅️ Voltar"):
    st.switch_page("main.py")
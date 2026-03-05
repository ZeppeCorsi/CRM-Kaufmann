import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, date, timedelta
import calendar
import os
import time as t_module

# --- 1. TRAVA DE SEGURANÇA ---
if 'logged_in' not in st.session_state or not st.session_state.logged_in:
    st.error("🚨 Acesso negado. Por favor, faça login na página inicial.")
    st.stop()

# --- 2. FUNÇÕES DE CONEXÃO ---
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

def carregar_aba(nome_aba):
    try:
        sh = conectar_google_sheets()
        worksheet = sh.worksheet(nome_aba)
        data = worksheet.get_all_records()
        if not data: return pd.DataFrame()
        df = pd.DataFrame(data)
        df = df.fillna("").astype(str)
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        if nome_aba == "Agendamentos" and "DATA" in df.columns:
            df['DATA_DT'] = pd.to_datetime(df['DATA'], errors='coerce', dayfirst=True).dt.date
            df = df.sort_values(by=["DATA_DT", "HORA"])
        return df
    except Exception as e:
        st.error(f"Erro ao carregar {nome_aba}: {e}")
        return pd.DataFrame()

def atualizar_visita_gs(indice_original, novo_orc_resp, data_follow, novos_detalhes):
    try:
        sh = conectar_google_sheets()
        worksheet = sh.worksheet("Agendamentos")
        # +2 porque o Python começa em 0 e o Sheets tem cabeçalho (linha 1)
        linha = int(indice_original) + 2
        worksheet.update_acell(f'G{linha}', "SIM") # Coluna G: Realizada
        worksheet.update_acell(f'L{linha}', data_follow.strftime("%d/%m/%Y")) # Coluna L: Follow-up
        worksheet.update_acell(f'O{linha}', novo_orc_resp) # Coluna O: Novo Orçamento
        
        # Busca obs atual para anexar o resultado
        obs_atual = worksheet.acell(f'H{linha}').value or ""
        nova_obs = f"{obs_atual} | RESULTADO: {novos_detalhes}".strip(" | ")
        worksheet.update_acell(f'H{linha}', nova_obs)
        
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
    follow = st.date_input("Próxima Data de Follow-up", value=date.today() + timedelta(days=7))
    relato = st.text_area("Descreva o que foi tratado na visita:")
    
    if st.button("Gravar na Planilha"):
        if atualizar_visita_gs(idx, novo_orc, follow, relato):
            st.balloons()
            st.success("✅ Visita finalizada com sucesso!")
            t_module.sleep(1)
            st.rerun()

# --- 4. INTERFACE DO CALENDÁRIO ---
st.title("📅 Calendário de Visitas")

# Controle de Navegação de Meses
if 'mes_ref' not in st.session_state:
    st.session_state.mes_ref = date.today().replace(day=1)

col_nav1, col_nav2, col_nav3 = st.columns([1, 2, 1])
if col_nav1.button("⬅️ Mês Anterior"):
    st.session_state.mes_ref = (st.session_state.mes_ref - timedelta(days=1)).replace(day=1)
    st.rerun()

col_nav2.markdown(f"<h3 style='text-align: center;'>{st.session_state.mes_ref.strftime('%B %Y').upper()}</h3>", unsafe_allow_html=True)

if col_nav3.button("Próximo Mês ➡️"):
    st.session_state.mes_ref = (st.session_state.mes_ref + timedelta(days=32)).replace(day=1)
    st.rerun()

# Carregamento dos dados
df_ag = carregar_aba("Agendamentos")

# Renderização do Calendário (Grid)
cal = calendar.monthcalendar(st.session_state.mes_ref.year, st.session_state.mes_ref.month)
dias_semana = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
cols_h = st.columns(7)
for i, d in enumerate(dias_semana):
    cols_h[i].markdown(f"<p style='text-align:center;font-weight:bold;color:#555;'>{d}</p>", unsafe_allow_html=True)

for semana in cal:
    cols = st.columns(7)
    for i, dia in enumerate(semana):
        if dia == 0:
            continue
        with cols[i]:
            data_atual = date(st.session_state.mes_ref.year, st.session_state.mes_ref.month, dia)
            # Destaque para o dia de hoje
            cor_dia = "blue" if data_atual == date.today() else "black"
            st.markdown(f"<p style='text-align:left; font-weight:bold; color:{cor_dia};'>{dia}</p>", unsafe_allow_html=True)
            
            if not df_ag.empty and 'DATA_DT' in df_ag.columns:
                visitas_dia = df_ag[df_ag['DATA_DT'] == data_atual]
                
                for idx, v in visitas_dia.iterrows():
                    realizada = str(v.get('REALIZADA', '')).upper() == "SIM"
                    icone = "✅" if realizada else "📍"
                    cli = v.get('CLIENTE', 'N/A')
                    hora = v.get('HORA', '--:--')
                    
                    # Expander compacto para cada visita
                    label = f"{icone} {hora} {cli[:10]}"
                    with st.expander(label):
                        st.write(f"**👤 Cliente:** {cli}")
                        st.write(f"**📞 Contato:** {v.get('NOME DO CONTATO', 'N/A')}")
                        st.write(f"**💰 Valor:** R$ {v.get('VALOR TOTAL', '0,00')}")
                        st.caption(f"📝 {v.get('DETALHES  DA VISITA', '')}")
                        
                        if not realizada:
                            if st.button("Finalizar", key=f"fin_{idx}"):
                                popup_finalizar_visita(idx, cli)

st.divider()
if st.button("⬅️ Voltar para o Início"):
    st.switch_page("main.py")
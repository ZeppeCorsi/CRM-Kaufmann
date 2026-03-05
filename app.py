import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, date, timedelta, time
import calendar
import urllib.parse
import os
import time as t_module # Importado para o delay do sucesso

# --- CONFIGURAÇÃO GOOGLE SHEETS ---
def conectar_google_sheets():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    
    # Transforma o segredo em dicionário editável (Evita erro de 'item assignment')
    creds_info = st.secrets["gcp_service_account"].to_dict()
    
    if "private_key" in creds_info:
        # Limpeza robusta da chave para evitar erro <Response [200]>
        pk = creds_info["private_key"].strip().strip('"').strip("'")
        pk = pk.replace("\\n", "\n")
        creds_info["private_key"] = pk
        
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_info, scope)
    client = gspread.authorize(creds)
    
    # Seu ID atualizado
    ID_PLANILHA = "1FI41GZwLTglXT4SAXIEyY53AXuheQg7gb_3pz9pWer0"
    
    return client.open_by_key(ID_PLANILHA)

def carregar_aba(nome_aba):
    try:
        sh = conectar_google_sheets()
        worksheet = sh.worksheet(nome_aba)
        data = worksheet.get_all_records()
        if not data: return pd.DataFrame()
        
        df = pd.DataFrame(data)
        # Limpeza de dados: remove 'nan' e espaços extras
        df = df.fillna("").astype(str).replace("nan", "")
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        if nome_aba == "Agendamentos" and "DATA" in df.columns:
            df['DATA_DT'] = pd.to_datetime(df['DATA'], errors='coerce', dayfirst=True).dt.date
            if "HORA" in df.columns:
                df = df.sort_values(by=["DATA_DT", "HORA"])
        return df
    except Exception as e:
        st.error(f"Erro ao carregar {nome_aba}: {e}")
        return pd.DataFrame()

def salvar_e_atualizar(novo_df):
    try:
        sh = conectar_google_sheets()
        worksheet = sh.worksheet("Agendamentos")
        novo_df = novo_df.astype(str)
        valores = novo_df.values.tolist()
        worksheet.append_rows(valores)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Erro ao salvar: {e}")
        return False

def atualizar_visita_gs(indice_original, novo_orc_resp, data_follow, novos_detalhes):
    try:
        sh = conectar_google_sheets()
        worksheet = sh.worksheet("Agendamentos")
        linha = int(indice_original) + 2
        
        # Atualiza colunas (Certifique-se que G, L e O batem com sua planilha)
        worksheet.update_acell(f'G{linha}', "SIM") # REALIZADA
        worksheet.update_acell(f'L{linha}', data_follow.strftime("%d/%m/%Y")) # FOLLOW-UP
        worksheet.update_acell(f'O{linha}', novo_orc_resp) # NOVO ORÇAMENTO
        
        obs_atual = worksheet.acell(f'H{linha}').value or ""
        nova_obs = f"{obs_atual} | RESULTADO: {novos_detalhes}".strip(" | ")
        worksheet.update_acell(f'H{linha}', nova_obs)
        
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Erro ao finalizar: {e}")
        return False

# --- FUNÇÃO OUTLOOK ---
def gerar_link_outlook(cliente, data, hora, finalidade, detalhes, contato):
    assunto = f"Agendamento de Visita - {cliente} ({data})"
    corpo = f"Olá,\n\nCliente: {cliente}\nData: {data} às {hora}\nFinalidade: {finalidade}\nContato: {contato}\n\nDetalhes:\n{detalhes}"
    return f"mailto:?subject={urllib.parse.quote(assunto)}&body={urllib.parse.quote(corpo)}"

# --- INTERFACE ---
st.set_page_config(page_title="Kaufmann CRM", layout="wide")
LOGO_PATH = "Logo_Kaufmann.jpg"
if os.path.exists(LOGO_PATH): st.sidebar.image(LOGO_PATH, use_container_width=True)

# --- POPUP FINALIZAR ---
@st.dialog("Finalizar Visita Realizada")
def popup_finalizar_visita(idx, cliente):
    st.write(f"Resultado da visita: **{cliente}**")
    novo_orc = st.radio("Gerou um NOVO Orçamento?", ["SIM", "NAO"], horizontal=True)
    follow = st.date_input("Próxima Data de Follow-up", value=date.today() + timedelta(days=7))
    relato = st.text_area("Descreva o que foi tratado:")
    if st.button("Gravar na Planilha"):
        if not relato: 
            st.warning("Conteúdo obrigatório.")
        elif atualizar_visita_gs(idx, novo_orc, follow, relato):
            st.balloons() # Celebração ao finalizar
            st.success("✅ Visita finalizada e atualizada!")
            t_module.sleep(2)
            st.rerun()

menu = st.sidebar.radio("Menu", ["📅 Calendário Comercial", "➕ Novo Agendamento"])

if menu == "📅 Calendário Comercial":
    st.title("📅 Calendário de Visitas")
    df_ag = carregar_aba("Agendamentos")
    
    if 'mes_ref' not in st.session_state: 
        st.session_state.mes_ref = date.today().replace(day=1)
    
    c1, c2, c3 = st.columns([1, 2, 1])
    if c1.button("⬅️ Anterior"):
        st.session_state.mes_ref = (st.session_state.mes_ref - timedelta(days=1)).replace(day=1)
        st.rerun()
    c2.markdown(f"<h3 style='text-align: center;'>{st.session_state.mes_ref.strftime('%B %Y').upper()}</h3>", unsafe_allow_html=True)
    if c3.button("Próximo ➡️"):
        st.session_state.mes_ref = (st.session_state.mes_ref + timedelta(days=32)).replace(day=1)
        st.rerun()

    cal = calendar.monthcalendar(st.session_state.mes_ref.year, st.session_state.mes_ref.month)
    cols_h = st.columns(7)
    dias_semana = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]
    for i, d in enumerate(dias_semana): 
        cols_h[i].markdown(f"<p style='text-align:center;font-weight:bold;'>{d}</p>", unsafe_allow_html=True)

    for semana in cal:
        cols = st.columns(7)
        for i, dia in enumerate(semana):
            if dia == 0: continue
            with cols[i]:
                data_atual = date(st.session_state.mes_ref.year, st.session_state.mes_ref.month, dia)
                st.markdown(f"**{dia}**")
                if not df_ag.empty and 'DATA_DT' in df_ag.columns:
                    visitas = df_ag[df_ag['DATA_DT'] == data_atual]
                    for idx, v in visitas.iterrows():
                        realizada = str(v.get('REALIZADA')).upper() == "SIM"
                        cor = "✅" if realizada else "📍"
                        with st.expander(f"{cor} {v.get('HORA','')} {v['CLIENTE'][:10]}"):
                            st.write(f"**Cliente:** {v['CLIENTE']}")
                            st.caption(f"{v.get('DETALHES  DA VISITA')}")
                            
                            link_mail = gerar_link_outlook(v['CLIENTE'], v['DATA'], v.get('HORA',''), v.get('FINALIDADE DA VISITA',''), v.get('DETALHES  DA VISITA',''), v.get('CONTATO',''))
                            st.markdown(f'<a href="{link_mail}"><button style="width:100%;">📧 Outlook</button></a>', unsafe_allow_html=True)
                            if not realizada:
                                if st.button("Finalizar", key=f"fin_{idx}"): 
                                    popup_finalizar_visita(idx, v['CLIENTE'])

elif menu == "➕ Novo Agendamento":
    st.title("➕ Novo Agendamento")
    df_orc = carregar_aba("Orcamentos Gerais")
    df_para = carregar_aba("Para_Agendar")

    col1, col2, col3 = st.columns([2, 2, 1])
    data_v = col1.date_input("Data", value=date.today())
    finalidade = col2.selectbox("Finalidade", ["ORCAMENTO", "PROSPECCAO", "POS VENDA"])
    hora_v = col3.time_input("Hora", value=time(9, 0))

    cliente_f = ""; vlr_f = ""; orc_num = ""; endereco_f = ""; cep_f = ""; contato_f = ""; tel_f = ""; email_f = ""

    if finalidade == "ORCAMENTO" and not df_para.empty:
        lista_cli = sorted(df_para.iloc[:, 0].dropna().astype(str).unique().tolist())
        cliente_f = st.selectbox("Selecione o Cliente", options=lista_cli)
        if cliente_f:
            busca = df_para[df_para.iloc[:, 0].astype(str).str.upper() == str(cliente_f).upper()]
            if not busca.empty:
                endereco_f = busca.iloc[0, 1]
                cep_f = busca.iloc[0, 3] if len(busca.columns) > 3 else ""
                st.success(f"📍 Endereço: {endereco_f}")
            if not df_orc.empty:
                m = df_orc[df_orc["CLIENTE"].astype(str).str.upper() == str(cliente_f).upper()]
                if not m.empty:
                    vlr_f = m.iloc[0].get("VALOR TOTAL", 0)
                    orc_num = m.iloc[0].get("ORCAMENTO", "")
                    st.info(f"💰 Orçamento: {orc_num}")

    elif finalidade == "PROSPECCAO":
        c1, c2 = st.columns(2)
        cliente_f = c1.text_input("Cliente")
        contato_f = c2.text_input("Contato")
        c3, c4 = st.columns(2)
        tel_f = c3.text_input("Telefone")
        email_f = c4.text_input("E-mail")

    with st.form("f_final"):
        if finalidade != "PROSPECCAO": 
            contato_f = st.text_input("Contato", value=contato_f)
        obs = st.text_area("Notas Adicionais")
        
        if st.form_submit_button("CONFIRMAR AGENDAMENTO"):
            if not cliente_f: 
                st.error("Cliente obrigatório")
            else:
                detalhes = f"END: {endereco_f} | CEP: {cep_f} | TEL: {tel_f} | EMAIL: {email_f} | {obs}"
                novo = pd.DataFrame([{
                    "DATA": data_v.strftime("%d/%m/%Y"), 
                    "HORA": hora_v.strftime("%H:%M"),
                    "FINALIDADE DA VISITA": finalidade, 
                    "CLIENTE": cliente_f,
                    "ORCAMENTO": orc_num, 
                    "VALOR TOTAL": vlr_f, 
                    "REALIZADA": "NAO", 
                    "DETALHES  DA VISITA": detalhes, 
                    "CONTATO": contato_f,
                    "DATA FOLLOW": "", 
                    "NOVO ORCAMENTO": ""
                }])
                
                if salvar_e_atualizar(novo):
                    st.balloons() # Efeito de balões
                    st.toast(f"Visita para {cliente_f} agendada!", icon="✅")
                    st.success("✅ Agendamento registrado com sucesso!")
                    t_module.sleep(2) # Pausa para ver o sucesso
                    st.rerun()
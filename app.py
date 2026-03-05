import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, date, timedelta, time
import calendar
import urllib.parse
import os
import time as t_module

# --- CONFIGURAÇÃO GOOGLE SHEETS ---
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
        # Limpeza: Colunas em MAIÚSCULO e sem espaços extras nas pontas
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        if nome_aba == "Agendamentos" and "DATA" in df.columns:
            df['DATA_DT'] = pd.to_datetime(df['DATA'], errors='coerce', dayfirst=True).dt.date
            df = df.sort_values(by=["DATA_DT", "HORA"])
        return df
    except Exception as e:
        st.error(f"Erro ao carregar {nome_aba}: {e}")
        return pd.DataFrame()

# ... (Funções salvar_e_atualizar e atualizar_visita_gs permanecem iguais) ...

def salvar_e_atualizar(novo_df):
    try:
        sh = conectar_google_sheets()
        worksheet = sh.worksheet("Agendamentos")
        valores = novo_df.astype(str).values.tolist()
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
        worksheet.update_acell(f'G{linha}', "SIM") 
        worksheet.update_acell(f'L{linha}', data_follow.strftime("%d/%m/%Y"))
        worksheet.update_acell(f'O{linha}', novo_orc_resp)
        obs_atual = worksheet.acell(f'H{linha}').value or ""
        nova_obs = f"{obs_atual} | RESULTADO: {novos_detalhes}".strip(" | ")
        worksheet.update_acell(f'H{linha}', nova_obs)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Erro ao finalizar: {e}")
        return False

def gerar_link_outlook(cliente, data, hora, finalidade, detalhes, contato):
    assunto = f"Agendamento - {cliente} ({data})"
    corpo = f"Cliente: {cliente}\nData: {data} às {hora}\nFinalidade: {finalidade}\nContato: {contato}\n\nDetalhes:\n{detalhes}"
    return f"mailto:?subject={urllib.parse.quote(assunto)}&body={urllib.parse.quote(corpo)}"

# --- INTERFACE ---
st.set_page_config(page_title="Kaufmann CRM", layout="wide")
LOGO_PATH = "Logo_Kaufmann.jpg"
if os.path.exists(LOGO_PATH): st.sidebar.image(LOGO_PATH, use_container_width=True)

@st.dialog("Finalizar Visita")
def popup_finalizar_visita(idx, cliente):
    st.write(f"Resultado da visita: **{cliente}**")
    novo_orc = st.radio("Gerou um NOVO Orçamento?", ["SIM", "NAO"], horizontal=True)
    follow = st.date_input("Próxima Data de Follow-up", value=date.today() + timedelta(days=7))
    relato = st.text_area("O que foi tratado?")
    if st.button("Gravar na Planilha"):
        if atualizar_visita_gs(idx, novo_orc, follow, relato):
            st.balloons()
            st.success("✅ Atualizado!")
            t_module.sleep(1); st.rerun()

menu = st.sidebar.radio("Menu", ["📅 Calendário Comercial", "➕ Novo Agendamento"])

if menu == "📅 Calendário Comercial":
    st.title("📅 Calendário de Visitas")
    df_ag = carregar_aba("Agendamentos")
    
    # ... (Lógica de navegação de meses permanece igual) ...
    if 'mes_ref' not in st.session_state: st.session_state.mes_ref = date.today().replace(day=1)
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
    for i, d in enumerate(["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]): 
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
                        realizada = str(v.get('REALIZADA', '')).upper() == "SIM"
                        cor = "✅" if realizada else "📍"
                        
                        # MAPEAMENTO DE COLUNAS (Ajustado para o que você informou)
                        cli = v.get('CLIENTE', 'N/A')
                        # Tentamos pegar NOME DO CONTATO ou CONTATO (o que existir)
                        con = v.get('NOME DO CONTATO') or v.get('CONTATO', 'N/A')
                        fin = v.get('FINALIDADE', 'N/A')
                        vlr = v.get('VALOR TOTAL', '0,00')
                        hora = v.get('HORA', '--:--')

                        with st.expander(f"{cor} {hora} {cli[:10]}"):
                            st.markdown(f"**👤 Cliente:** {cli}")
                            st.markdown(f"**📞 Contato:** {con}")
                            st.markdown(f"**🎯 Finalidade:** {fin}")
                            st.markdown(f"**💰 Valor:** R$ {vlr}")
                            
                            st.divider()
                            st.caption(f"📝 {v.get('DETALHES  DA VISITA','')}")
                            
                            if st.button("Finalizar", key=f"fin_{idx}"): 
                                popup_finalizar_visita(idx, cli)

elif menu == "➕ Novo Agendamento":
    st.title("➕ Novo Agendamento")
    df_para = carregar_aba("Para_Agendar")
    df_orc_gerais = carregar_aba("Orcamentos Gerais")

    col1, col2, col3 = st.columns([2, 2, 1])
    data_v = col1.date_input("Data", value=date.today())
    finalidade = col2.selectbox("Finalidade", ["ORCAMENTO", "PROSPECCAO", "POS VENDA"])
    hora_v = col3.time_input("Hora", value=time(9, 0))

    cliente_f = ""; vlr_f = "0,00"; orc_num = ""; endereco_f = ""

    if not df_para.empty:
        # Pega a coluna CLIENTE independente de onde ela esteja
        col_cli = [c for c in df_para.columns if 'CLIENTE' in c]
        lista_cli = sorted(df_para[col_cli[0]].unique().tolist()) if col_cli else []
        cliente_f = st.selectbox("Selecione o Cliente", options=lista_cli)
        
        if cliente_f:
            dados_cli = df_para[df_para[col_cli[0]] == cliente_f]
            if not dados_cli.empty:
                vlr_f = dados_cli.iloc[0].get("VLR TOTAL", "0,00")
                endereco_f = dados_cli.iloc[0].get("ENDEREÇO", "")
            
            if not df_orc_gerais.empty:
                dados_orc = df_orc_gerais[df_orc_gerais["CLIENTE"] == cliente_f]
                if not dados_orc.empty:
                    orc_num = dados_orc.iloc[0].get("ORCAMENTO", "")

    with st.form("f_final"):
        contato_f = st.text_input("Nome do Contato") # Aqui você preenche
        obs = st.text_area("Observações Adicionais")
        
        if st.form_submit_button("CONFIRMAR AGENDAMENTO"):
            detalhes = f"Endereço: {endereco_f} | Obs: {obs}"
            # AS COLUNAS ABAIXO DEVEM SER IGUAIS À SUA PLANILHA AGENDAMENTOS
            novo = pd.DataFrame([{
                "DATA": data_v.strftime("%d/%m/%Y"), 
                "HORA": hora_v.strftime("%H:%M"),
                "FINALIDADE": finalidade, # Salva como 'Finalidade'
                "CLIENTE": cliente_f,
                "ORCAMENTO": orc_num, 
                "VALOR TOTAL": vlr_f, 
                "REALIZADA": "NAO", 
                "DETALHES  DA VISITA": detalhes, 
                "NOME DO CONTATO": contato_f, # Salva como 'Nome do contato'
                "DATA FOLLOW": "", 
                "NOVO ORCAMENTO": ""
            }])
            if salvar_e_atualizar(novo):
                st.balloons(); st.success("Agendado!"); t_module.sleep(1); st.rerun()
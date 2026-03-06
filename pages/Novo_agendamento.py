import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, date, time, timedelta
import os
import time as t_module

# --- 1. TRAVA DE SEGURANÇA ---
if 'logged_in' not in st.session_state or not st.session_state.logged_in:
    st.error("🚨 Acesso negado. Por favor, faça login na página inicial.")
    st.stop()

# --- 2. FUNÇÕES DE APOIO ---
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
        return df
    except Exception as e:
        st.error(f"Erro ao carregar {nome_aba}: {e}")
        return pd.DataFrame()

def formatar_br(valor_str):
    try:
        limpo = str(valor_str).replace("R$", "").replace(".", "").replace(",", ".").strip()
        valor_float = float(limpo)
        return f"{valor_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return valor_str

def salvar_agendamento(novo_df):
    try:
        sh = conectar_google_sheets()
        worksheet = sh.worksheet("Agendamentos")
        # Converte para lista de listas para o append_rows
        valores = novo_df.astype(str).values.tolist()
        worksheet.append_rows(valores)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Erro ao salvar: {e}")
        return False

# --- 3. INTERFACE DA PÁGINA ---
st.title("➕ Novo Agendamento")
st.caption(f"Logado como: **{st.session_state.user_data['nome']}** | Perfil: **{st.session_state.user_data['perfil']}**")

with st.spinner('Sincronizando com Google Sheets...'):
    df_para = carregar_aba("Para_Agendar")
    df_orc_gerais = carregar_aba("Orcamentos Gerais")

if df_para.empty:
    st.warning("⚠️ Nenhuma lista de clientes encontrada em 'Para_Agendar'.")
else:
    col1, col2, col3 = st.columns([2, 2, 1])
    
    # AJUSTE: Formato de exibição brasileiro no calendário do sistema
    data_v = col1.date_input("Data da Visita", value=date.today(), format="DD/MM/YYYY")
    
    finalidade = col2.selectbox("Finalidade", ["ORCAMENTO", "PROSPECCAO", "POS VENDA", "REAGENDADA"])
    hora_v = col3.time_input("Hora da Visita", value=time(9, 0))

    col_cli = [c for c in df_para.columns if 'CLIENTE' in c]
    lista_cli = sorted(df_para[col_cli[0]].unique().tolist()) if col_cli else []
    cliente_f = st.selectbox("Selecione o Cliente", options=[""] + lista_cli)
    
    vlr_f = "0,00"
    orc_num = "Não localizado"
    endereco_f = ""

    if cliente_f:
        dados_cli = df_para[df_para[col_cli[0]].str.strip() == cliente_f.strip()]
        if not dados_cli.empty:
            vlr_raw = dados_cli.iloc[0].get("VLR TOTAL", "0,00")
            vlr_f = formatar_br(vlr_raw)
            endereco_f = dados_cli.iloc[0].get("ENDEREÇO", "")
        
        if not df_orc_gerais.empty:
            dados_orc = df_orc_gerais[df_orc_gerais["CLIENTE"].str.strip() == cliente_f.strip()]
            if not dados_orc.empty:
                orc_num = dados_orc.iloc[0].get("ORCAMENTO", "Não localizado")

        c_vlr, c_orc = st.columns(2)
        c_vlr.metric("💰 Valor Estimado", f"R$ {vlr_f}")
        c_orc.metric("📄 Orçamento Atual", orc_num)
        
        if endereco_f:
            st.info(f"📍 **Endereço Base:** {endereco_f}")

    with st.form("form_agendamento"):
        contato_f = st.text_input("Nome do Contato / Responsável")
        obs = st.text_area("Observações Adicionais (Detalhes da Visita)")
        
        enviar = st.form_submit_button("🚀 CONFIRMAR AGENDAMENTO")
        
        if enviar:
            if not cliente_f:
                st.error("❌ Erro: Selecione um cliente antes de confirmar.")
            else:
                detalhes_completos = f"Endereço: {endereco_f} | Obs: {obs}"
                
                # AJUSTE: Capturando Horário de Brasília (UTC-3)
                # O servidor do Streamlit geralmente usa UTC (00:00).
                agora_utc = datetime.now()
                agora_br = agora_utc - timedelta(hours=3)
                
                # Criando o registro com Auditoria de Usuário e Data/Hora da inclusão
                novo_registro = pd.DataFrame([{
                    "DATA": data_v.strftime("%d/%m/%Y"), 
                    "HORA": hora_v.strftime("%H:%M"),
                    "FINALIDADE": finalidade,
                    "CLIENTE": cliente_f,
                    "ORCAMENTO": orc_num if orc_num != "Não localizado" else "", 
                    "VALOR TOTAL": vlr_f, 
                    "REALIZADA": "NAO", 
                    "DETALHES DA VISITA": detalhes_completos, 
                    "NOME DO CONTATO": contato_f,
                    "USUARIO_INCLUSAO": st.session_state.user_data['nome'], 
                    "DATA_HORA_LOG": agora_br.strftime("%d/%m/%Y %H:%M:%S")
                }])
                
                if salvar_agendamento(novo_registro):
                    st.balloons()
                    st.success("✅ Agendamento gravado com sucesso!")
                    t_module.sleep(2)
                    st.rerun()

st.divider()
if st.button("⬅️ Voltar para o Início"):
    st.switch_page("main.py")
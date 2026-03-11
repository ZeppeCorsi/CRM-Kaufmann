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

# --- 3. INTERFACE DA PÁGINA ---
st.title("➕ Novo Agendamento")
st.caption(f"Logado como: **{st.session_state.user_data['nome']}** | Perfil: **{st.session_state.user_data['perfil']}**")

with st.spinner('Sincronizando com Google Sheets...'):
    df_para = carregar_aba("Para_Agendar")
    df_orc_gerais = carregar_aba("Orcamentos Gerais")

# Configurações Iniciais de Colunas
col1, col2, col3 = st.columns([2, 2, 1])
data_v = col1.date_input("Data da Visita", value=date.today(), format="DD/MM/YYYY")
finalidade = col2.selectbox("Finalidade", ["ORCAMENTO", "PROSPECCAO", "POS VENDA", "REAGENDADA"])
hora_v = col3.time_input("Hora da Visita", value=time(9, 0))

# Variáveis de dados
cliente_f = ""
vlr_f = "0,00"
orc_num = ""
endereco_f = ""
contato_f = ""

# --- LÓGICA CONDICIONAL POR FINALIDADE ---
if finalidade == "PROSPECCAO":
    st.info("💡 **Modo Prospecção:** Digite os dados manualmente abaixo.")
    cliente_f = st.text_input("Nome da Empresa (Novo Cliente)")
    contato_f = st.text_input("Nome do Contato / Responsável")
    endereco_f = st.text_input("Endereço ou Referência")
    forma_contato = st.text_input("Telefone ou E-mail de Contato")
    vlr_f = "0,00"
    orc_num = ""
else:
    if df_para.empty:
        st.warning("⚠️ Lista de clientes vazia.")
    else:
        col_cli = [c for c in df_para.columns if 'CLIENTE' in c]
        lista_cli = sorted(df_para[col_cli[0]].unique().tolist()) if col_cli else []
        cliente_selecionado = st.selectbox("Selecione o Cliente", options=[""] + lista_cli)
        
        if cliente_selecionado:
            cliente_f = cliente_selecionado
            dados_cli = df_para[df_para[col_cli[0]].str.strip() == cliente_f.strip()]
            if not dados_cli.empty:
                vlr_raw = dados_cli.iloc[0].get("VLR TOTAL", "0,00")
                vlr_f = formatar_br(vlr_raw)
                endereco_f = dados_cli.iloc[0].get("ENDEREÇO", "")
            
            if not df_orc_gerais.empty:
                dados_orc = df_orc_gerais[df_orc_gerais["CLIENTE"].str.strip() == cliente_f.strip()]
                if not dados_orc.empty:
                    orc_num = dados_orc.iloc[0].get("ORCAMENTO", "")

            c_vlr, c_orc = st.columns(2)
            c_vlr.metric("💰 Valor Estimado", f"R$ {vlr_f}")
            c_orc.metric("📄 Orçamento Atual", orc_num if orc_num else "S/N")
            
            if endereco_f:
                st.info(f"📍 **Endereço Base:** {endereco_f}")
            
            contato_f = st.text_input("Nome do Contato / Responsável")

# --- FORMULÁRIO DE FINALIZAÇÃO ---
with st.form("form_agendamento"):
    obs = st.text_area("Observações Adicionais (Detalhes da Visita)")
    enviar = st.form_submit_button("🚀 CONFIRMAR AGENDAMENTO")
    
    if enviar:
        if not cliente_f:
            st.error("❌ Erro: Informe o nome do cliente antes de confirmar.")
        else:
            # Auditoria de Horário (Brasília)
            agora_br = (datetime.now() - timedelta(hours=3)).strftime("%d/%m/%Y %H:%M:%S")
            
            # Unificação do endereço com as observações para a coluna H
            detalhes_para_planilha = f"Endereço: {endereco_f} | Obs: {obs}" if finalidade == "PROSPECCAO" else f"{endereco_f} | Obs: {obs}"

            # --- MONTAGEM DA LINHA COMPLETA (ORDEM RÍGIDA A até N) ---
            # Cada posição nesta lista é uma coluna física no Google Sheets
            nova_linha = [
                data_v.strftime("%d/%m/%Y"),           # A - DATA
                hora_v.strftime("%H:%M"),              # B - HORARIO
                finalidade,                            # C - FINALIDADE
                cliente_f,                             # D - CLIENTE
                orc_num,                               # E - ORCAMENTO
                vlr_f,                                 # F - VALOR TOTAL
                "NAO",                                 # G - GEROU ORC
                detalhes_para_planilha,                # H - ENDERECO (DETALHES DA VISITA)
                contato_f,                             # I - CONTATO
                st.session_state.user_data['nome'],    # J - Usuario
                agora_br,                              # K - INCLUSAO
                "NAO",                                 # L - REALIZADA
                "",                                    # M - FOLLOW-UP
                ""                                     # N - RELATO FINAL (Notas do Popup)
            ]
            
            try:
                sh = conectar_google_sheets()
                worksheet = sh.worksheet("Agendamentos")
                
                # O uso do append_row (singular) com a lista direta impede o deslocamento de colunas
                worksheet.append_row(nova_linha, value_input_option='RAW')
                
                st.balloons()
                st.success("✅ Agendamento gravado com sucesso na Coluna A!")
                st.cache_data.clear()
                t_module.sleep(2)
                st.rerun()
            except Exception as e:
                st.error(f"Erro ao salvar: {e}")

st.divider()
if st.button("⬅️ Voltar para o Início"):
    st.switch_page("main.py")
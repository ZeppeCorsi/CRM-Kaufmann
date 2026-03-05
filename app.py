import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, date, timedelta, time
import calendar
import urllib.parse
import os
import time as t_module

# --- 1. CONFIGURAÇÕES INICIAIS E SEGURANÇA ---
st.set_page_config(page_title="Kaufmann CRM", layout="wide")
LOGO_PATH = "Logo_Kaufmann.jpg"

# Tabela de acessos (No futuro, podemos puxar isso da planilha)
USUARIOS = {
    "adm_kaufmann": {"senha": "123", "perfil": "ADM", "nome": "Administrador"},
    "coord_vendas": {"senha": "456", "perfil": "COORD", "nome": "Coordenador"},
    "vendedor_01": {"senha": "789", "perfil": "VEND", "nome": "Vendedor 1"},
    "vendedor_02": {"senha": "000", "perfil": "VEND", "nome": "Vendedor 2"},
}

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

def formatar_br(valor_str):
    try:
        limpo = str(valor_str).replace("R$", "").replace(".", "").replace(",", ".").strip()
        valor_float = float(limpo)
        return f"{valor_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return valor_str

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

# --- 3. LÓGICA DE LOGIN ---
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    # --- TELA DE ACESSO ---
    c1, c2, c3 = st.columns([1, 1.5, 1])
    with c2:
        if os.path.exists(LOGO_PATH): 
            st.image(LOGO_PATH, use_container_width=True) #
        st.markdown("<h2 style='text-align: center;'>Acesso ao Sistema CRM</h2>", unsafe_allow_html=True)
        usuario = st.text_input("Usuário")
        senha = st.text_input("Senha", type="password")
        if st.button("ENTRAR", use_container_width=True):
            if usuario in USUARIOS and USUARIOS[usuario]["senha"] == senha:
                st.session_state.logged_in = True
                st.session_state.user_data = USUARIOS[usuario]
                st.rerun()
            else:
                st.error("Usuário ou senha incorretos.")

else:
    # --- 4. INTERFACE PRINCIPAL (USUÁRIO LOGADO) ---
    user = st.session_state.user_data
    if os.path.exists(LOGO_PATH): 
        st.sidebar.image(LOGO_PATH, use_container_width=True)
    
    st.sidebar.write(f"👤 **{user['nome']}**")
    st.sidebar.caption(f"Perfil: {user['perfil']}")

    # Menu dinâmico por perfil
    opcoes_menu = ["📅 Calendário Comercial", "➕ Novo Agendamento"]
    if user['perfil'] in ["ADM", "COORD"]:
        opcoes_menu.append("📊 Relatório Gerencial")
    
    menu = st.sidebar.radio("Navegação", opcoes_menu)
    
    if st.sidebar.button("Sair/Logout"):
        st.session_state.logged_in = False
        st.rerun()

    # --- PÁGINA: CALENDÁRIO COMERCIAL ---
    if menu == "📅 Calendário Comercial":
        st.title("📅 Calendário de Visitas")
        df_ag = carregar_aba("Agendamentos")
        
        # Lógica de navegação de meses (Sessão)
        if 'mes_ref' not in st.session_state: st.session_state.mes_ref = date.today().replace(day=1)
        c1, c2, c3 = st.columns([1, 2, 1])
        if c1.button("⬅️ Anterior"):
            st.session_state.mes_ref = (st.session_state.mes_ref - timedelta(days=1)).replace(day=1)
            st.rerun()
        c2.markdown(f"<h3 style='text-align: center;'>{st.session_state.mes_ref.strftime('%B %Y').upper()}</h3>", unsafe_allow_html=True)
        if c3.button("Próximo ➡️"):
            st.session_state.mes_ref = (st.session_state.mes_ref + timedelta(days=32)).replace(day=1)
            st.rerun()

        # Renderização do Calendário
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
                            realizada = str(v.get('REALIZADA', '')).upper() == "SIM"
                            cor = "✅" if realizada else "📍"
                            cli = v.get('CLIENTE', 'N/A')
                            with st.expander(f"{cor} {v.get('HORA', '--:--')} {cli[:8]}..."):
                                st.write(f"**Cliente:** {cli}")
                                st.write(f"**Valor:** R$ {v.get('VALOR TOTAL', '0,00')}")
                                if st.button("Finalizar", key=f"btn_{idx}"):
                                    # Chamada do seu popup_finalizar_visita (pode ser adicionado aqui)
                                    pass

    # --- PÁGINA: NOVO AGENDAMENTO ---
    elif menu == "➕ Novo Agendamento":
        st.title("➕ Novo Agendamento")
        df_para = carregar_aba("Para_Agendar")
        df_orc_gerais = carregar_aba("Orcamentos Gerais")

        col1, col2, col3 = st.columns([2, 2, 1])
        data_v = col1.date_input("Data", value=date.today())
        finalidade = col2.selectbox("Finalidade", ["ORCAMENTO", "PROSPECCAO", "POS VENDA"])
        hora_v = col3.time_input("Hora", value=time(9, 0))

        cliente_f = ""; vlr_f = "0,00"; orc_num = "Não localizado"; endereco_f = ""

        if not df_para.empty:
            col_cli = [c for c in df_para.columns if 'CLIENTE' in c]
            lista_cli = sorted(df_para[col_cli[0]].unique().tolist()) if col_cli else []
            cliente_f = st.selectbox("Selecione o Cliente", options=[""] + lista_cli)
            
            if cliente_f:
                dados_cli = df_para[df_para[col_cli[0]].str.strip() == cliente_f.strip()]
                if not dados_cli.empty:
                    vlr_f = formatar_br(dados_cli.iloc[0].get("VLR TOTAL", "0,00")) # Valor com ponto e vírgula
                    endereco_f = dados_cli.iloc[0].get("ENDEREÇO", "")
                
                if not df_orc_gerais.empty:
                    dados_orc = df_orc_gerais[df_orc_gerais["CLIENTE"].str.strip() == cliente_f.strip()]
                    if not dados_orc.empty:
                        orc_num = dados_orc.iloc[0].get("ORCAMENTO", "Não localizado")
                
                # Cartões de Métricas
                c_vlr, c_orc = st.columns(2)
                c_vlr.metric("💰 Valor Estimado", f"R$ {vlr_f}")
                c_orc.metric("📄 Orçamento Vinculado", orc_num)

        with st.form("f_final"):
            contato_f = st.text_input("Nome do Contato")
            obs = st.text_area("Observações Adicionais")
            if st.form_submit_button("CONFIRMAR AGENDAMENTO"):
                if not cliente_f:
                    st.error("Selecione um cliente!")
                else:
                    detalhes = f"Endereço: {endereco_f} | Obs: {obs}"
                    novo = pd.DataFrame([{
                        "DATA": data_v.strftime("%d/%m/%Y"), 
                        "HORA": hora_v.strftime("%H:%M"),
                        "FINALIDADE": finalidade,
                        "CLIENTE": cliente_f,
                        "ORCAMENTO": orc_num if orc_num != "Não localizado" else "", 
                        "VALOR TOTAL": vlr_f, 
                        "REALIZADA": "NAO", 
                        "DETALHES  DA VISITA": detalhes, 
                        "NOME DO CONTATO": contato_f,
                        "DATA FOLLOW": "", 
                        "NOVO ORCAMENTO": ""
                    }])
                    if salvar_e_atualizar(novo):
                        st.balloons(); st.success("Agendado!"); t_module.sleep(1); st.rerun()

    # --- PÁGINA: RELATÓRIO GERENCIAL ---
    elif menu == "📊 Relatório Gerencial":
        st.title("📊 Relatório Gerencial")
        st.info("Painel de indicadores disponível apenas para ADM e Coordenadores.")
        # Espaço para os futuros gráficos
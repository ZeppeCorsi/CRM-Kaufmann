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
        
        # Normaliza nomes das colunas para evitar erros de busca
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        if nome_aba == "Agendamentos" and "DATA" in df.columns:
            # Converte data para objeto datetime do Python (formato brasileiro)
            df['DATA_DT'] = pd.to_datetime(df['DATA'], errors='coerce', dayfirst=True).dt.date
            
            # Ordenação segura: verifica se HORA existe antes de ordenar
            colunas_ordem = ["DATA_DT"]
            if "HORA" in df.columns:
                colunas_ordem.append("HORA")
            
            df = df.sort_values(by=colunas_ordem)
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
        
        # Mapeamento de colunas atualizado (conforme ordem do Novo Agendamento)
        # G=REALIZADA, J=NOME DO CONTATO, K=USUARIO_INCLUSAO, L=DATA_HORA_LOG
        # Vamos usar índices numéricos para garantir precisão
        worksheet.update_cell(linha, 7, "SIM") # Coluna G: Realizada
        
        # Busca obs atual (Coluna H) para anexar o resultado
        obs_atual = worksheet.cell(linha, 8).value or ""
        nova_obs = f"{obs_atual} | RESULTADO: {novos_detalhes}".strip(" | ")
        worksheet.update_cell(linha, 8, nova_obs) # Coluna H: Detalhes/Obs
        
        # Se houver colunas para Follow-up e Novo Orçamento na sua planilha,
        # ajuste os números abaixo (ex: se forem as colunas 12 e 13)
        # worksheet.update_cell(linha, 12, data_follow.strftime("%d/%m/%Y")) 
        
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

# --- 4. INTERFACE DO CALENDÁRIO ---
st.title("📅 Calendário de Visitas")

# Controle de Navegação de Meses
if 'mes_ref' not in st.session_state:
    st.session_state.mes_ref = date.today().replace(day=1)

col_nav1, col_nav2, col_nav3 = st.columns([1, 2, 1])
if col_nav1.button("⬅️ Mês Anterior"):
    st.session_state.mes_ref = (st.session_state.mes_ref - timedelta(days=1)).replace(day=1)
    st.rerun()

# Nome do mês em português
meses_pt = {
    1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO", 4: "ABRIL", 5: "MAIO", 6: "JUNHO",
    7: "JULHO", 8: "AGOSTO", 9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO"
}
nome_mes = meses_pt[st.session_state.mes_ref.month]
col_nav2.markdown(f"<h3 style='text-align: center;'>{nome_mes} {st.session_state.mes_ref.year}</h3>", unsafe_allow_html=True)

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
                        # Tratamento para erro de espaço duplo no nome da coluna
                        detalhes = v.get('DETALHES DA VISITA', v.get('DETALHES  DA VISITA', ''))
                        st.caption(f"📝 {detalhes}")
                        
                        if not realizada:
                            if st.button("Finalizar", key=f"fin_{idx}"):
                                popup_finalizar_visita(idx, cli)

st.divider()
if st.button("⬅️ Voltar para o Início"):
    st.switch_page("main.py")
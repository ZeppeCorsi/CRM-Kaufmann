import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta

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

def carregar_dados_relatorio():
    try:
        sh = conectar_google_sheets()
        worksheet = sh.worksheet("Agendamentos")
        data = worksheet.get_all_records()
        if not data: 
            return pd.DataFrame()
        df = pd.DataFrame(data)
        # Padroniza nomes de colunas para garantir leitura da Coluna N
        df.columns = [str(c).strip().upper() for c in df.columns]
        return df
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return pd.DataFrame()

# --- 3. PROCESSAMENTO DE DADOS ---
st.set_page_config(page_title="Relatório de Visitas", layout="wide")
st.title("📊 Relatório de Visitas e Prospecções")

df_rel = carregar_dados_relatorio()

if df_rel.empty:
    st.warning("⚠️ Nenhum dado encontrado na aba Agendamentos.")
else:
    # Tratamento de Datas e Valores
    df_rel['DATA_DT'] = pd.to_datetime(df_rel['DATA'], dayfirst=True, errors='coerce')
    
    def limpar_valor(v):
        try:
            return float(str(v).replace('R$', '').replace('.', '').replace(',', '.').strip())
        except:
            return 0.0
    df_rel['VALOR_NUM'] = df_rel['VALOR TOTAL'].apply(limpar_valor)

    # --- 4. FILTROS LATERAIS ---
    st.sidebar.header("Filtros")
    data_inicio = st.sidebar.date_input("Início", value=datetime.now() - timedelta(days=30))
    data_fim = st.sidebar.date_input("Fim", value=datetime.now())
    
    mask = (df_rel['DATA_DT'].dt.date >= data_inicio) & (df_rel['DATA_DT'].dt.date <= data_fim)
    df_filtrado = df_rel.loc[mask].copy()

    # Visitas Reais (Ignora REAGENDADA nas métricas principais)
    df_visitas_reais = df_filtrado[df_filtrado['FINALIDADE'] != "REAGENDADA"].copy()

    # --- 5. MÉTRICAS ---
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Visitas", len(df_visitas_reais))
    m2.metric("Prospecções", len(df_visitas_reais[df_visitas_reais['FINALIDADE'] == "PROSPECCAO"]))
    m3.metric("Orçamentos", len(df_visitas_reais[df_visitas_reais['FINALIDADE'] == "ORCAMENTO"]))
    m4.metric("Pós-Venda", len(df_visitas_reais[df_visitas_reais['FINALIDADE'] == "POS VENDA"]))

    st.divider()

    # --- 6. DETALHES DAS VISITAS (COLUNA N) ---
    st.subheader("📝 Relatos e Notas das Visitas")
    # Filtramos apenas visitas que possuem algum relato na Coluna N
    # O nome exato da coluna na sua planilha é "DETALHES DA VISITA" (Coluna N)
    col_detalhes = "DETALHES DA VISITA" 
    
    if col_detalhes in df_filtrado.columns:
        df_com_relato = df_filtrado[df_filtrado[col_detalhes] != ""].copy()
        if not df_com_relato.empty:
            for _, row in df_com_relato.iterrows():
                temp = str(row.get('TEMPERATURA', '')).strip()
                icone_temp = {"QUENTE": "🔥", "MORNO": "🌡️", "FRIO": "🧊"}.get(temp.upper(), "")
                label_temp = f" | {icone_temp} {temp}" if temp else ""
                with st.expander(f"📌 {row['DATA']} - {row['CLIENTE']} ({row['FINALIDADE']}){label_temp}"):
                    st.write(f"**Vendedor:** {row['USUARIO']}")
                    st.write(f"**Relato Final:** {row[col_detalhes]}")
                    if temp:
                        st.write(f"**Temperatura:** {icone_temp} {temp}")
                    if 'DATA FOLLOW' in row and row['DATA FOLLOW']:
                        st.info(f"📅 Próximo Follow-up: {row['DATA FOLLOW']}")
        else:
            st.info("Nenhuma visita com relato preenchido no período.")
    else:
        st.error(f"Coluna '{col_detalhes}' não encontrada na planilha.")

    # --- 7. ANÁLISES GRÁFICAS ---
    st.divider()
    c_graf1, c_graf2 = st.columns(2)

    with c_graf1:
        st.write("### 🏆 Top Clientes")
        st.bar_chart(df_visitas_reais['CLIENTE'].value_counts().head(5))

    with c_graf2:
        st.write("### 📍 Top Regiões")
        def extrair_regiao(end):
            partes = str(end).split(',')
            return partes[-1].strip().upper() if len(partes) > 1 else "N/A"
        df_visitas_reais['REGIAO'] = df_visitas_reais['ENDERECO'].apply(extrair_regiao)
        st.table(df_visitas_reais['REGIAO'].value_counts().head(5))

    # --- 8. TABELA COMPLETA ---
    st.write("### 📋 Tabela Geral")
    cols_view = ['DATA', 'CLIENTE', 'FINALIDADE', 'VALOR TOTAL', 'USUARIO', 'REALIZADA', 'TEMPERATURA', col_detalhes]
    cols_existentes = [c for c in cols_view if c in df_filtrado.columns]
    st.dataframe(df_filtrado[cols_existentes], use_container_width=True)

st.sidebar.divider()
if st.sidebar.button("🔄 Atualizar"):
    st.cache_data.clear()
    st.rerun()
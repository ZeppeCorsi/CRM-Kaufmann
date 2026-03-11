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
        # Padroniza nomes de colunas para evitar erros de digitação
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
    st.warning("⚠️ Nenhum dado encontrado na aba Agendamentos ou aba inexistente.")
else:
    # Tratamento de Datas
    df_rel['DATA_DT'] = pd.to_datetime(df_rel['DATA'], dayfirst=True, errors='coerce')
    
    # Tratamento de Valores Numéricos
    def limpar_valor(v):
        try:
            return float(str(v).replace('R$', '').replace('.', '').replace(',', '.').strip())
        except:
            return 0.0
    df_rel['VALOR_NUM'] = df_rel['VALOR TOTAL'].apply(limpar_valor)

    # --- 4. FILTROS LATERAIS ---
    st.sidebar.header("Filtros de Análise")
    data_inicio = st.sidebar.date_input("Data Início", value=datetime.now() - timedelta(days=30))
    data_fim = st.sidebar.date_input("Data Fim", value=datetime.now())
    
    # Aplicar Filtro de Período
    mask = (df_rel['DATA_DT'].dt.date >= data_inicio) & (df_rel['DATA_DT'].dt.date <= data_fim)
    df_filtrado = df_rel.loc[mask].copy()

    # Separar Visitas Reais (Ignora Reagendamentos para estatística)
    df_visitas_reais = df_filtrado[df_filtrado['FINALIDADE'] != "REAGENDADA"].copy()

    # --- 5. EXIBIÇÃO DE MÉTRICAS ---
    st.subheader(f"📌 Resumo: {data_inicio.strftime('%d/%m/%Y')} até {data_fim.strftime('%d/%m/%Y')}")
    
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total de Visitas", len(df_visitas_reais))
    m2.metric("Prospecções", len(df_visitas_reais[df_visitas_reais['FINALIDADE'] == "PROSPECCAO"]))
    m3.metric("Orçamentos", len(df_visitas_reais[df_visitas_reais['FINALIDADE'] == "ORCAMENTO"]))
    m4.metric("Pós-Venda", len(df_visitas_reais[df_visitas_reais['FINALIDADE'] == "POS VENDA"]))

    st.divider()

    # --- 6. ANÁLISES GRÁFICAS ---
    col_graf1, col_graf2 = st.columns(2)

    with col_graf1:
        st.write("### 🏆 Clientes mais Visitados")
        if not df_visitas_reais.empty:
            top_clientes = df_visitas_reais['CLIENTE'].value_counts().head(10)
            st.bar_chart(top_clientes)
        else:
            st.info("Sem dados para o gráfico.")

    with col_graf2:
        st.write("### 💰 Cliente com Maior Valor Acumulado")
        if not df_visitas_reais.empty:
            # Agrupa por cliente e soma o valor total
            vendas_cliente = df_visitas_reais.groupby('CLIENTE')['VALOR_NUM'].sum().sort_values(ascending=False).head(1)
            if not vendas_cliente.empty:
                cliente_topo = vendas_cliente.index[0]
                valor_topo = vendas_cliente.values[0]
                st.info(f"O destaque é **{cliente_topo}**")
                st.metric("Total Estimado", f"R$ {valor_topo:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        else:
            st.info("Sem dados financeiros.")

    # --- 7. ANÁLISE POR REGIÃO ---
    st.write("### 📍 Regiões (Extraído do Endereço)")
    if not df_visitas_reais.empty:
        # Tenta pegar a última parte do endereço (geralmente a cidade)
        def extrair_regiao(end):
            partes = str(end).split(',')
            if len(partes) > 1:
                return partes[-1].strip().upper()
            return str(end)[:25].upper() # Fallback para os primeiros caracteres se não houver vírgula

        df_visitas_reais['REGIAO'] = df_visitas_reais['ENDERECO'].apply(extrair_regiao)
        regioes = df_visitas_reais['REGIAO'].value_counts().head(5)
        st.table(regioes)

    # --- 8. PRÉVIA DA PLANILHA ---
    st.divider()
    st.write("### 📋 Prévia dos Dados Filtrados")
    # Colunas que existem na sua planilha
    colunas_exibicao = ['DATA', 'HORARIO', 'FINALIDADE', 'CLIENTE', 'VALOR TOTAL', 'USUARIO']
    # Filtrar apenas colunas que existem no DF para evitar erro
    colunas_df = [c for c in colunas_exibicao if c in df_filtrado.columns]
    st.dataframe(df_filtrado[colunas_df], use_container_width=True)

st.sidebar.divider()
if st.sidebar.button("🔄 Atualizar Relatório"):
    st.cache_data.clear()
    st.rerun()
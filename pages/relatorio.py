import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

# --- 1. CONFIGURAÇÃO E CARREGAMENTO ---
st.title("📊 Relatório de Visitas e Prospecções")

def carregar_dados_agendamentos():
    try:
        # Reutilizando sua função de conexão
        from main import carregar_aba 
        df = carregar_aba("Agendamentos")
        if df.empty:
            return pd.DataFrame()
        
        # Converter coluna de DATA para formato datetime para filtrar por período
        df['DATA_DT'] = pd.to_datetime(df['DATA'], dayfirst=True, errors='coerce')
        # Limpar valores de VALOR TOTAL para cálculo numérico
        df['VALOR_NUM'] = df['VALOR TOTAL'].str.replace('.', '').str.replace(',', '.').astype(float)
        return df
    except Exception as e:
        st.error(f"Erro ao processar dados para o relatório: {e}")
        return pd.DataFrame()

df_rel = carregar_dados_agendamentos()

if df_rel.empty:
    st.warning("Nenhum dado encontrado na aba Agendamentos.")
    st.stop()

# --- 2. FILTROS DE PERÍODO ---
col_f1, col_f2 = st.columns(2)
data_inicio = col_f1.date_input("Data Início", value=datetime.now() - timedelta(days=30))
data_fim = col_f2.date_input("Data Fim", value=datetime.now())

# Filtrar o DataFrame pelo período selecionado
mask = (df_rel['DATA_DT'].dt.date >= data_inicio) & (df_rel['DATA_DT'].dt.date <= data_fim)
df_filtrado = df_rel.loc[mask]

# --- 3. MÉTRICAS PRINCIPAIS ---
st.subheader("📌 Resumo do Período")

# Quantidade total ignorando REAGENDADA
df_visitas_reais = df_filtrado[df_filtrado['FINALIDADE'] != "REAGENDADA"]
total_visitas = len(df_visitas_reais)

c1, c2, c3, c4 = st.columns(4)
c1.metric("Total Visitas", total_visitas)
c2.metric("Prospecções", len(df_visitas_reais[df_visitas_reais['FINALIDADE'] == "PROSPECCAO"]))
c3.metric("Orçamentos", len(df_visitas_reais[df_visitas_reais['FINALIDADE'] == "ORCAMENTO"]))
c4.metric("Pós-Venda", len(df_visitas_reais[df_visitas_reais['FINALIDADE'] == "POS VENDA"]))

st.divider()

# --- 4. ANÁLISES DETALHADAS ---
col_an1, col_an2 = st.columns(2)

with col_an1:
    st.write("### 🏆 Clientes mais Visitados")
    top_clientes = df_visitas_reais['CLIENTE'].value_counts().head(5)
    st.bar_chart(top_clientes)

with col_an2:
    st.write("### 💰 Maior Valor Estimado")
    # Busca o cliente com maior valor na coluna VALOR TOTAL
    top_valor = df_visitas_reais.sort_values(by='VALOR_NUM', ascending=False).head(1)
    if not top_valor.empty:
        st.info(f"**{top_valor['CLIENTE'].values[0]}**")
        st.metric("Valor", f"R$ {top_valor['VALOR TOTAL'].values[0]}")

# --- 5. ANÁLISE POR REGIÃO (EXTRAÇÃO DO ENDEREÇO) ---
st.write("### 📍 Visitas por Cidade/Região")
# Tentativa simples: pega a primeira palavra após o endereço ou assume que a cidade está no texto
def extrair_cidade(texto):
    # Exemplo: Se o endereço for "RUA SALVADOR... SÃO PAULO", tentamos capturar padrões comuns
    texto = str(texto).upper()
    cidades = ["SÃO PAULO", "GUARULHOS", "CAMPINAS", "SANTOS", "SÃO BERNARDO"] # Adicione as cidades da sua região
    for cidade in cidades:
        if cidade in texto:
            return cidade
    return "OUTRAS / NÃO IDENTIF."

df_visitas_reais['REGIAO'] = df_visitas_reais['ENDERECO'].apply(extrair_cidade)
regiao_counts = df_visitas_reais['REGIAO'].value_counts()
st.dataframe(regiao_counts, use_container_width=True)

# --- 6. PRÉVIA DOS DADOS ---
st.divider()
st.write("### 📋 Detalhes das Visitas no Período")
st.dataframe(df_filtrado[['DATA', 'HORARIO', 'FINALIDADE', 'CLIENTE', 'VALOR TOTAL', 'USUARIO']], use_container_width=True)
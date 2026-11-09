import streamlit as st
import pandas as pd
import plotly.express as px

# Configuração da página do painel
st.set_page_config(page_title="Dashboard de Vendas", layout="wide")
st.title("📊 Painel de Vendas e Devoluções")

# 1. Lendo os dados da sua planilha existente
# O pandas lê as abas formatadas facilmente
planilha_coleta = r'y:\ADMINISTRATIVO\Vanessa - Vendas do E-Commerce\COLETA E-COMMERCE 2026.xlsx'

df_vendas = pd.read_excel(planilha_coleta, sheet_name="Vendas")
df_devolucoes = pd.read_excel(planilha_coleta, sheet_name="Devoluções")

# 2. Criando blocos de métricas no topo (KPIs)
col1, col2, col3 = st.columns(3)
with col1:
    st.metric("Faturamento Total", f"R$ {df_vendas['VALOR'].sum():,.2f}")
with col2:
    st.metric("Total Devoluções", f"R$ {df_devolucoes['VALOR'].sum():,.2f}")
with col3:
    st.metric("Qtd. Itens Vendidos", f"{df_vendas['QTD'].sum()}")

# 3. Criando gráficos interativos
st.subheader("Análise de Desempenho")
grafico_linha = px.line(df_vendas, x="DATA", y="VALOR", title="Evolução das Vendas no Tempo")
st.plotly_chart(grafico_linha, use_container_width=True)

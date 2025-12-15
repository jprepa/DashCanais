import streamlit as st
import pandas as pd
import plotly.express as px

# Configuração da Página
st.set_page_config(page_title="Painel Comercial - Prevision", layout="wide")

# Título e Subtítulo
st.title("📊 Painel Comercial DDD - Sienge Plataforma")
st.markdown("---")

# --- 1. CARREGAMENTO DOS DADOS ---
# Aqui você vai colocar o caminho do seu arquivo baixado na sua máquina por enquanto
# Dica: O Python tem dificuldade de ler direto do link do SharePoint sem autenticação complexa.
# O ideal é você baixar o Excel ou mapear a pasta do OneDrive no seu PC.
@st.cache_data
def carregar_dados():
    arquivo = 'Ideia Nih.xlsm' # Nome do seu arquivo
    
    # Carregando as abas (baseado nos nomes que vi nos prints)
    # Usei skiprows e usecols se precisar limpar cabeçalhos, mas aqui vou direto
    df_vendas = pd.read_excel(arquivo, sheet_name='Vendas')
    df_churn = pd.read_excel(arquivo, sheet_name='Churn')
    
    return df_vendas, df_churn

try:
    # Tenta carregar (coloquei try caso o arquivo não esteja na pasta)
    df_vendas, df_churn = carregar_dados()
    
    # --- 2. BARRA LATERAL (FILTROS) ---
    st.sidebar.header("Filtros")
    
    # Filtro de Ano (assumindo que tem coluna 'Ano')
    anos_disponiveis = df_vendas['Ano'].unique()
    ano_selecionado = st.sidebar.selectbox("Selecione o Ano", anos_disponiveis)
    
    # Filtrar os dataframes
    vendas_filtrado = df_vendas[df_vendas['Ano'] == ano_selecionado]
    
    # --- 3. KPIs (CARTÕES) ---
    st.subheader(f"Resumo Executivo - {ano_selecionado}")
    
    col1, col2, col3, col4 = st.columns(4)
    
    # Exemplo de cálculo (ajuste os nomes das colunas conforme seu Excel real)
    # Vi no print que tem uma coluna 'NMRR Novas Vendas' na aba Painel, 
    # mas na aba Vendas deve ter o valor individual.
    
    total_vendas = vendas_filtrado['NMRR Adicionado'].sum() # Ajuste o nome da coluna
    qtd_vendas = vendas_filtrado['Nº Contrato'].count()
    ticket_medio = total_vendas / qtd_vendas if qtd_vendas > 0 else 0
    
    col1.metric("NMRR Total", f"R$ {total_vendas:,.2f}")
    col2.metric("Qtd. Vendas", f"{qtd_vendas}")
    col3.metric("Ticket Médio", f"R$ {ticket_medio:,.2f}")
    
    # --- 4. GRÁFICOS ---
    
    c1, c2 = st.columns(2)
    
    with c1:
        st.markdown("### Vendas por Mês")
        # Agrupando por mês
        vendas_mensal = vendas_filtrado.groupby('Mês')['NMRR Adicionado'].sum().reset_index()
        fig_vendas = px.bar(vendas_mensal, x='Mês', y='NMRR Adicionado', text_auto=True, color_discrete_sequence=['#C00000'])
        st.plotly_chart(fig_vendas, use_container_width=True)
        
    with c2:
        st.markdown("### Mix de Produtos/Porte")
        # Vi que tem coluna 'Porte'
        fig_pizza = px.pie(vendas_filtrado, names='Porte', values='NMRR Adicionado', hole=0.5)
        st.plotly_chart(fig_pizza, use_container_width=True)

except FileNotFoundError:
    st.error("Arquivo 'Ideia Nih.xlsm' não encontrado na pasta. Coloque o arquivo junto com o script.")
except Exception as e:
    st.error(f"Ocorreu um erro: {e}")
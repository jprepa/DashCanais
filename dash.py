import streamlit as st
import pandas as pd
import plotly.express as px
from unicodedata import normalize

st.set_page_config(page_title="Análise Canais", layout="wide")
st.title("📊 Análise Canais (Filtro DDD + Canal)")
st.markdown("---")

# --- 1. MAPEAMENTOS ---

# Canais por DDD (Sua regra de negócio)
MAPA_CANAIS = {
    'TCP/GSN': [31, 32, 33, 34, 35, 36, 37, 38],
    'NPU': [47, 48, 49],
    'TISEN/TWA/CONSTRUSOFT': [21, 22, 24],
    'MR SOLER': [27, 28],
    'BRPRO': [41, 42],
    'NG7': [83, 84, 87, 71, 73, 74, 75, 77, 82],
    'Controller': [12],
    'Excelencia': [51, 53, 54, 55],
    'Foco': [98, 99, 86, 89],
    'Gavazzi': [85, 88],
    'Gescon': [11, 13, 19],
    'INACX': [14, 15],
    'JC MANANCIAL': [91, 93, 94],
    'JC RAMIREZ': [92, 95, 96, 97],
    'Pontara': [43, 44, 45, 46, 67],
    'PSA': [61, 62, 63, 64, 65, 66, 69],
    'PZR': [16, 17, 18],
    'Ronaldo Chagas': [68],
    'Vercosa': [79],
    'Softplan': [0, 1]
}

# Meses (Nome <-> Número)
MAPA_MESES = {
    1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
    5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
    9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
}
MAPA_MESES_INV = {v: k for k, v in MAPA_MESES.items()}

# --- FUNÇÕES AUXILIARES ---
def definir_canal(ddd_valor):
    try:
        ddd_int = int(ddd_valor)
    except:
        return "Sem DDD"
    for canal, lista_ddds in MAPA_CANAIS.items():
        if ddd_int in lista_ddds: return canal
    return "Outros (Sem Canal)"

def normalizar_cabecalho(texto):
    if not isinstance(texto, str): return str(texto)
    return normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII').lower().strip()

def encontrar_coluna(df, palavras_chave):
    colunas_norm = {col: normalizar_cabecalho(col) for col in df.columns}
    for palavra in palavras_chave:
        palavra_norm = normalizar_cabecalho(palavra)
        for col_orig, col_n in colunas_norm.items():
            if col_n == palavra_norm: return col_orig
        for col_orig, col_n in colunas_norm.items():
            if palavra_norm in col_n: return col_orig
    return None

def limpar_dinheiro(valor):
    if pd.isna(valor): return 0.0
    if isinstance(valor, (int, float)): return float(valor)
    s = str(valor).strip().upper().replace('R$', '').replace(' ', '').replace('\xa0', '')
    if not s: return 0.0
    s = s.replace('.', '').replace(',', '.')
    try: return float(s)
    except: return 0.0

def formatar_reais(valor):
    # Formatação padrao R$
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def formatar_numero_decimal(valor):
    # Formatação sem R$ (para Mercado Total)
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def formatar_qtd(valor):
    return f"{int(valor)}"

def formatar_porcentagem(valor):
    if pd.isna(valor): return "0.0%"
    if valor > 1.0: return f"{valor:.1f}%" # Caso venha como 12.5
    return f"{valor * 100:.1f}%" # Caso venha como 0.125

def extrair_ddd(valor):
    s = str(valor)
    nums = ''.join(filter(str.isdigit, s))
    if len(nums) >= 2: return int(nums[:2])
    return 0

# --- CONFIGURAÇÃO DE COLUNAS ---
CONFIG_BUSCA = {
    'Vendas': {'valor': ['MRR Adicionado', 'NMRR Novas Vendas', 'Valor', 'MRR'], 'data': ['Mês/Ano', 'Data', 'Data Venda', 'Mes', 'Ano']},
    'Aditivos': {'valor': ['MRR Adicionado', 'NMRR Adicionado', 'Valor Aditivo', 'Valor'], 'data': ['Mês/Ano', 'Data', 'Data Aditivo', 'Mes', 'Ano']},
    'SQL': {'data': ['Mês/Ano', 'Data', 'Data SQL', 'Mes', 'Ano']},
    'Churn': {'valor': ['MRR Perdido', 'MRR', 'Valor'], 'data': ['Mês/Ano', 'Data', 'Data Cancelamento', 'Mes', 'Ano']},
    'Reduções': {
        'valor': ['MRR Contrato', 'MRR Reduzido', 'Valor Redução', 'Valor'], 
        'data': ['Mês/Ano', 'Data', 'Data Redução', 'Mes', 'Ano'],
        'lifetime': ['Lifetime', 'Lifetime Meses', 'Tempo de Casa', 'Meses']
    },
    'Share': {
        'mercado_total': ['Mercado Total', 'Nº Mercado Total'],
        'mercado_pond': ['Mercado ponderado', 'Nº Mercado Ponderado'],
        'empresas': ['Empresas clientes', 'Qtd Clientes', 'Empresas'],
        'share_pond': ['Share ponderado', 'Média de Share Ponderado'],
        'sqls': ['SQLs', 'Nº SQLs'],
        'vendas': ['Vendas', 'Nº Vendas'],
        'conv': ['SQL x Vendas', 'Conversão']
    }
}

# --- PROCESSAMENTO ---
uploaded_file = st.file_uploader("Upload da planilha (.xlsx)", type=['xlsx', 'xlsm'])

dfs = {}
filtros_ano = set()
filtros_mes_nums = set()
# Conjunto para guardar pares unicos (DDD, Canal) encontrados
ddds_encontrados_info = set() 
ddds_sem_canal = set()

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        
        for aba_alvo, config in CONFIG_BUSCA.items():
            nome_aba_real = next((a for a in xls.sheet_names if normalizar_cabecalho(a) == normalizar_cabecalho(aba_alvo)), None)
            
            if nome_aba_real:
                df = pd.read_excel(xls, sheet_name=nome_aba_real)
                clean = pd.DataFrame()
                
                # Porte
                c_porte = encontrar_coluna(df, ['Porte', 'Classificação'])
                clean['Porte'] = df[c_porte].astype(str).str.strip().str.upper() if c_porte else 'ND'

                # Data (Exceto Share)
                if aba_alvo != 'Share':
                    c_ano = encontrar_coluna(df, ['Ano', 'Year'])
                    c_mes = encontrar_coluna(df, ['Mês', 'Mes', 'Month'])
                    
                    if c_ano and c_mes:
                        clean['Ano'] = pd.to_numeric(df[c_ano], errors='coerce').fillna(0).astype(int)
                        clean['Mês'] = pd.to_numeric(df[c_mes], errors='coerce').fillna(0).astype(int)
                    else:
                        c_data = encontrar_coluna(df, config.get('data', []))
                        if c_data:
                            datas = pd.to_datetime(df[c_data], errors='coerce')
                            clean['Ano'] = datas.dt.year.fillna(0).astype(int)
                            clean['Mês'] = datas.dt.month.fillna(0).astype(int)
                        else:
                            clean['Ano'] = 0
                            clean['Mês'] = 0
                    
                    filtros_ano.update(clean[clean['Ano'] > 2000]['Ano'].unique())
                    filtros_mes_nums.update(clean[clean['Mês'] > 0]['Mês'].unique())

                # DDD e Canal
                c_ddd = encontrar_coluna(df, ['DDD', 'Ddd', 'Telefone', 'Celular'])
                if c_ddd:
                    clean['DDD_Num'] = df[c_ddd].apply(extrair_ddd)
                    clean['Canal'] = clean['DDD_Num'].apply(definir_canal)
                    
                    # Guardar info para o filtro (DDD e Canal)
                    # Apenas DDDs válidos (>0)
                    pares_unicos = clean[clean['DDD_Num'] > 0][['DDD_Num', 'Canal']].drop_duplicates().values
                    for d, c in pares_unicos:
                        ddds_encontrados_info.add((int(d), c))

                    ddds_perdidos = clean[clean['Canal'] == 'Outros (Sem Canal)']['DDD_Num'].unique()
                    ddds_sem_canal.update(ddds_perdidos)
                else:
                    clean['Canal'] = 'Sem Info'
                    clean['DDD_Num'] = 0

                # Valores e Colunas Especiais
                if aba_alvo == 'Share':
                    clean['Qtd'] = 0
                    for key, keywords in config.items():
                        c_encontrada = encontrar_coluna(df, keywords)
                        if c_encontrada:
                            if 'mercado' in key: clean[key] = df[c_encontrada].apply(limpar_dinheiro)
                            else: clean[key] = pd.to_numeric(df[c_encontrada], errors='coerce').fillna(0)
                        else: clean[key] = 0.0
                else:
                    # Valor
                    if 'valor' in config:
                        c_valor = encontrar_coluna(df, config['valor'])
                        clean['Valor'] = df[c_valor].apply(limpar_dinheiro) if c_valor else 0.0
                    else: clean['Valor'] = 0.0
                    
                    # Lifetime (Reduções)
                    if 'lifetime' in config:
                        c_life = encontrar_coluna(df, config['lifetime'])
                        clean['Lifetime'] = pd.to_numeric(df[c_life], errors='coerce').fillna(0) if c_life else 0.0
                    
                    clean['Qtd'] = 1 

                dfs[aba_alvo] = clean

        # --- FILTROS ---
        st.sidebar.header("Filtros")
        
        # Ano
        lista_anos = sorted(list(filtros_ano), reverse=True)
        ano_sel = st.sidebar.selectbox("Ano", lista_anos) if lista_anos else None
        
        # Mês (Nomes)
        lista_meses_nums = sorted(list(filtros_mes_nums))
        lista_meses_nomes = [MAPA_MESES.get(m, str(m)) for m in lista_meses_nums]
        mes_sel_nomes = st.sidebar.multiselect("Mês", lista_meses_nomes, default=lista_meses_nomes)
        mes_sel_nums = [MAPA_MESES_INV.get(m) for m in mes_sel_nomes if m in MAPA_MESES_INV]
        
        # --- FILTRO DDD + CANAL ---
        # Ordena por DDD
        lista_ddds_ordenada = sorted(list(ddds_encontrados_info), key=lambda x: x[0])
        # Cria strings formatadas "DD - Canal"
        opcoes_formatadas = [f"{d} - {c}" for d, c in lista_ddds_ordenada]
        
        # Multiselect mostrando "47 - NPU"
        ddd_sel_strings = st.sidebar.multiselect("DDD - Canal", options=opcoes_formatadas, default=opcoes_formatadas)
        
        # Extrai apenas os números dos DDDs selecionados para filtrar o DF
        # Ex: "47 - NPU" -> pega 47
        ddds_escolhidos = [int(s.split(' - ')[0]) for s in ddd_sel_strings]

        st.sidebar.markdown("---")
        if ddds_sem_canal:
            with st.sidebar.expander("⚠️ DDDs sem Canal", expanded=True):
                st.write(sorted(list(ddds_sem_canal)))

        def filtrar_df(df):
            if df is None or df.empty: return pd.DataFrame()
            mask = pd.Series(True, index=df.index)
            if ano_sel and 'Ano' in df.columns: mask &= (df['Ano'] == ano_sel) | (df['Ano'] == 0)
            if mes_sel_nums and 'Mês' in df.columns: mask &= (df['Mês'].isin(mes_sel_nums)) | (df['Mês'] == 0)
            
            # Filtra por DDD_Num se houver seleção
            if ddds_escolhidos and 'DDD_Num' in df.columns:
                 # Filtra se o DDD está na lista
                 # (Para tabelas que não têm DDD (ex: share as vezes?), assume-se 0 ou mantem)
                 # Aqui filtramos estrito pelos DDDs presentes na lista
                 mask &= df['DDD_Num'].isin(ddds_escolhidos)
                 
            return df[mask]

        # --- ABAS ---
        tab_tabelas, tab_graficos = st.tabs(["📋 Visão Tabela", "📈 Visão Gráfica"])

        # =======================================================
        # TAB 1: TABELAS (ONE PAGER)
        # =======================================================
        with tab_tabelas:
            
            def gerar_tabela_simples(nome_aba, titulo_qtd, titulo_valor):
                df = filtrar_df(dfs.get(nome_aba))
                if df.empty: return pd.DataFrame(columns=['Porte', titulo_qtd, titulo_valor])
                resumo = df.groupby('Porte')[['Qtd', 'Valor']].sum().reset_index()
                total_q, total_v = resumo['Qtd'].sum(), resumo['Valor'].sum()
                final = pd.concat([resumo, pd.DataFrame({'Porte': ['Total Geral'], 'Qtd': [total_q], 'Valor': [total_v]})], ignore_index=True)
                final[titulo_qtd] = final['Qtd'].apply(formatar_qtd)
                final[titulo_valor] = final['Valor'].apply(formatar_reais)
                return final[['Porte', titulo_qtd, titulo_valor]]

            def gerar_tabela_share():
                df = filtrar_df(dfs.get('Share'))
                if df.empty: return pd.DataFrame()
                cols = ['mercado_total', 'mercado_pond', 'empresas', 'sqls', 'vendas']
                resumo = df.groupby('Porte')[cols].sum().reset_index()
                total = pd.DataFrame([['Total Geral'] + [resumo[c].sum() for c in cols]], columns=['Porte'] + cols)
                final = pd.concat([resumo, total], ignore_index=True)
                
                final['Share Pond'] = (final['empresas'] / final['mercado_pond']).fillna(0)
                final['Conv'] = (final['vendas'] / final['sqls']).fillna(0)
                
                # Sem R$ no Mercado Total/Ponderado
                final['Mercado Total'] = final['mercado_total'].apply(formatar_numero_decimal)
                final['Mercado Pond.'] = final['mercado_pond'].apply(formatar_numero_decimal)
                
                final['Empresas'] = final['empresas'].apply(formatar_qtd)
                final['SQLs'] = final['sqls'].apply(formatar_qtd)
                final['Vendas'] = final['vendas'].apply(formatar_qtd)
                final['Share Pond.'] = final['Share Pond'].apply(formatar_porcentagem)
                final['% Conv'] = final['Conv'].apply(formatar_porcentagem)
                
                return final[['Porte', 'Mercado Total', 'Mercado Pond.', 'Empresas', 'Share Pond.', 'SQLs', 'Vendas', '% Conv']]

            def gerar_tabela_churn():
                df_churn = filtrar_df(dfs.get('Churn'))
                df_vendas = filtrar_df(dfs.get('Vendas'))
                if df_churn.empty: return pd.DataFrame(columns=['Porte', 'Nº Clientes', 'MRR Perdido', '% Vendas x Churn'])
                
                resumo = df_churn.groupby('Porte')[['Qtd', 'Valor']].sum().reset_index()
                tot_q, tot_v = resumo['Qtd'].sum(), resumo['Valor'].sum()
                resumo = pd.concat([resumo, pd.DataFrame({'Porte': ['Total Geral'], 'Qtd': [tot_q], 'Valor': [tot_v]})], ignore_index=True)
                
                if not df_vendas.empty:
                    v_grp = df_vendas.groupby('Porte')['Valor'].sum().reset_index()
                    t_vendas = v_grp['Valor'].sum()
                    v_grp = pd.concat([v_grp, pd.DataFrame({'Porte': ['Total Geral'], 'Valor': [t_vendas]})], ignore_index=True)
                    resumo = pd.merge(resumo, v_grp, on='Porte', how='left', suffixes=('', '_Vendas')).fillna(0)
                else: resumo['Valor_Vendas'] = 0
                
                resumo['Ratio'] = (resumo['Valor'] / resumo['Valor_Vendas']).fillna(0)
                resumo['Nº Clientes'] = resumo['Qtd'].apply(formatar_qtd)
                resumo['MRR Perdido'] = resumo['Valor'].apply(formatar_reais)
                resumo['% Vendas x Churn'] = resumo['Ratio'].apply(formatar_porcentagem)
                return resumo[['Porte', 'Nº Clientes', 'MRR Perdido', '% Vendas x Churn']]

            def gerar_tabela_reducoes():
                df = filtrar_df(dfs.get('Reduções'))
                if df.empty: return pd.DataFrame(columns=['Porte', 'Nº Reduções', 'MRR Reduzido', 'Média Lifetime'])
                resumo = df.groupby('Porte').agg({'Qtd': 'sum', 'Valor': 'sum', 'Lifetime': 'mean'}).reset_index()
                tot_q, tot_v = resumo['Qtd'].sum(), resumo['Valor'].sum()
                media = df['Lifetime'].mean() if not df.empty else 0
                resumo = pd.concat([resumo, pd.DataFrame({'Porte': ['Total Geral'], 'Qtd': [tot_q], 'Valor': [tot_v], 'Lifetime': [media]})], ignore_index=True)
                
                resumo['Nº Reduções'] = resumo['Qtd'].apply(formatar_qtd)
                resumo['MRR Reduzido'] = resumo['Valor'].apply(formatar_reais)
                resumo['Média Lifetime'] = resumo['Lifetime'].apply(lambda x: f"{x:.1f} meses")
                return resumo[['Porte', 'Nº Reduções', 'MRR Reduzido', 'Média Lifetime']]

            c1, c2 = st.columns(2)
            with c1:
                st.markdown("### 🟢 Novas Vendas")
                st.dataframe(gerar_tabela_simples('Vendas', 'Nº Vendas', 'NMRR Novas Vendas'), hide_index=True, use_container_width=True)
            with c2:
                st.markdown("### 🔵 Aditivos")
                st.dataframe(gerar_tabela_simples('Aditivos', 'Nº Clientes', 'NMRR Adicionado'), hide_index=True, use_container_width=True)

            st.markdown("---")
            c3, c4 = st.columns(2)
            with c3:
                st.markdown("### 📊 SQL x Novas Vendas")
                df_sql, df_ven = filtrar_df(dfs.get('SQL')), filtrar_df(dfs.get('Vendas'))
                if not df_sql.empty and not df_ven.empty:
                    s_sql = df_sql.groupby('Porte')['Qtd'].sum().reset_index().rename(columns={'Qtd': 'SQLs'})
                    s_ven = df_ven.groupby('Porte')['Qtd'].sum().reset_index().rename(columns={'Qtd': 'Vendas'})
                    m = pd.merge(s_sql, s_ven, on='Porte', how='outer').fillna(0)
                    ts, tv = m['SQLs'].sum(), m['Vendas'].sum()
                    m = pd.concat([m, pd.DataFrame({'Porte': ['Total Geral'], 'SQLs': [ts], 'Vendas': [tv]})])
                    m['% Conv'] = (m['Vendas'] / m['SQLs']).fillna(0).apply(formatar_porcentagem)
                    m['SQLs'] = m['SQLs'].apply(formatar_qtd)
                    st.dataframe(m[['Porte', 'SQLs', '% Conv']], hide_index=True, use_container_width=True)
            with c4:
                st.markdown("### 🌎 Share de Mercado")
                st.dataframe(gerar_tabela_share(), hide_index=True, use_container_width=True)

            st.markdown("---")
            c5, c6 = st.columns(2)
            with c5:
                st.markdown("### 🔴 Churn")
                st.dataframe(gerar_tabela_churn(), hide_index=True, use_container_width=True)
            with c6:
                st.markdown("### 📉 Reduções")
                st.dataframe(gerar_tabela_reducoes(), hide_index=True, use_container_width=True)

        # =======================================================
        # TAB 2: GRÁFICOS (VISUAL)
        # =======================================================
        with tab_graficos:
            st.markdown("### 📊 Comparativo Visual")
            col_sel, _ = st.columns([1, 3])
            with col_sel:
                agrupar_por = st.selectbox("Agrupar gráficos por:", ["Canal", "Porte"])
            
            def get_dados_grafico(nome_aba):
                df = filtrar_df(dfs.get(nome_aba))
                if df.empty or agrupar_por not in df.columns: return pd.DataFrame()
                return df.groupby(agrupar_por)[['Qtd', 'Valor']].sum().reset_index()

            g1, g2 = st.columns(2)
            
            # Gráfico Vendas
            dv = get_dados_grafico('Vendas')
            if not dv.empty:
                fig1 = px.bar(dv, x=agrupar_por, y='Valor', title=f"Vendas (R$) por {agrupar_por}", text_auto='.2s', color=agrupar_por)
                g1.plotly_chart(fig1, use_container_width=True)
            else: g1.info("Sem dados para Vendas")

            # Gráfico Vendas vs Churn
            dc = get_dados_grafico('Churn')
            if not dv.empty and not dc.empty:
                v, c = dv.copy(), dc.copy()
                v['Tipo'], c['Tipo'] = 'Vendas', 'Churn'
                combined = pd.concat([v, c])
                fig2 = px.bar(combined, x=agrupar_por, y='Valor', color='Tipo', barmode='group',
                              title=f"Vendas vs Churn (R$) por {agrupar_por}",
                              color_discrete_map={'Vendas': '#00CC96', 'Churn': '#EF553B'})
                g2.plotly_chart(fig2, use_container_width=True)
            else: g2.info("Dados insuficientes para Comparativo")

            st.markdown("---")
            
            # Gráfico Funil
            dsql = get_dados_grafico('SQL')
            if not dsql.empty and not dv.empty:
                s, v = dsql.copy(), dv.copy()
                s['Etapa'], s['Val'] = 'SQLs', s['Qtd']
                v['Etapa'], v['Val'] = 'Vendas', v['Qtd']
                funnel = pd.concat([s, v])
                fig3 = px.funnel(funnel, x='Val', y=agrupar_por, color='Etapa', title=f"Funil (Qtd) por {agrupar_por}")
                st.plotly_chart(fig3, use_container_width=True)

    except Exception as e:
        st.error(f"Erro ao processar: {e}")

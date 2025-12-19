# app.py - VERSÃO ATUALIZADA COM MATRIZES DE RESUMO POR CRITICIDADE E CURVA

import streamlit as st
import pandas as pd
import numpy as np
from scipy.stats import norm
import io

# ==============================
# CONFIGURAÇÃO DA PÁGINA
# ==============================
st.set_page_config(page_title="Estoque de Segurança - Revisão Periódica", layout="wide")
st.title("Calculadora de Estoque de Segurança")
st.markdown("**Sistema de Revisão Periódica (revisão mensal)** | Período de proteção = Lead Time + 1 mês")

st.info("""
**Fórmula utilizada:**

$$SS = z \\times \\sqrt{(LT + 1) \\times \\sigma_d^2 + d^2 \\times \\sigma_{LT}^2}$$

Onde z é baseado no nível de serviço configurado na matriz por Criticidade (X,Y,Z) e Curva (A,B,C).
""")

# ==============================
# UPLOAD E TEMPLATE
# ==============================
st.header("Upload de Dados")

# Colunas para o template (incluindo novas)
columns_template = [
    'Empresa', 'Classe', 'SKU', 'Descrição_do_Material',
    'Estoque inicial (unidades)', 'Lead Time médio (meses)',
    'Desvio Padrão LT (meses)', 'Estoque de Segurança (unidades)',
    'Plano de Demanda (un/mês)', 'Horizonte do PD (meses)',
    'Estoque em Trânsito inicial', 'Consumo Médio Mensal (un)',
    'Desvio Padrão Consumo (un)', 'melhor_distribuicao',
    'parametros', 'Valor Unit', 'Criticidade', 'Curva'
]

df_template = pd.DataFrame(columns=columns_template)
buffer_template = io.BytesIO()
with pd.ExcelWriter(buffer_template, engine='openpyxl') as writer:
    df_template.to_excel(writer, index=False, sheet_name='dados')
buffer_template.seek(0)

st.download_button(
    label="Baixar template Excel (preencha na sheet 'dados')",
    data=buffer_template,
    file_name="template_dados_skus.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.markdown("""
**Orientações para o template:**
- Use a sheet nomeada como **'dados'**.
- Preencha as colunas exatamente como no template.
- **Criticidade:** X, Y ou Z.
- **Curva:** A, B ou C.
- Colunas obrigatórias: 'Lead Time médio (meses)', 'Desvio Padrão LT (meses)', 
  'Consumo Médio Mensal (un)', 'Desvio Padrão Consumo (un)', 'Valor Unit', 
  'Estoque de Segurança (unidades)', 'Criticidade', 'Curva'.
""")

# Upload do arquivo
uploaded_file = st.file_uploader("Upload sua base de dados (Excel, sheet 'dados')", type="xlsx")

if uploaded_file is None:
    st.info("Faça upload do arquivo para prosseguir.")
    st.stop()

# Carregamento dinâmico
try:
    df = pd.read_excel(uploaded_file, sheet_name="dados")
except Exception as e:
    st.error(f"Erro ao carregar: {e}. Verifique a sheet 'dados'.")
    st.stop()

df.columns = df.columns.str.strip()
num_cols = [
    'Lead Time médio (meses)', 'Desvio Padrão LT (meses)',
    'Consumo Médio Mensal (un)', 'Desvio Padrão Consumo (un)',
    'Valor Unit', 'Estoque inicial (unidades)', 'Estoque de Segurança (unidades)'
]
str_cols = ['Criticidade', 'Curva', 'SKU']

for col in num_cols:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    else:
        st.error(f"Coluna obrigatória ausente: {col}")
        st.stop()

for col in str_cols:
    if col in df.columns:
        df[col] = df[col].astype(str).str.upper()  # Padronizar para maiúsculas
    else:
        st.error(f"Coluna obrigatória ausente: {col}")
        st.stop()

df = df.fillna(0)

# ==============================
# SIDEBAR - VISÃO GERAL
# ==============================
st.sidebar.header("Visão Geral dos Dados Carregados")
df['Valor_Estoque_Inicial'] = df['Estoque inicial (unidades)'] * df['Valor Unit']
valor_total_estoque = df['Valor_Estoque_Inicial'].sum()

st.sidebar.metric("Total de SKUs únicos", f"{df['SKU'].nunique():,}")
st.sidebar.metric("Total de empresas", df['Empresa'].nunique())
st.sidebar.metric("Valor total do estoque atual", f"R$ {valor_total_estoque:,.0f}")
st.sidebar.write("**Empresas:** " + ", ".join(sorted(df['Empresa'].unique())))

# ==============================
# CONFIGURAÇÃO DA MATRIZ DE NÍVEIS DE SERVIÇO
# ==============================
st.header("Matriz de Níveis de Serviço (%) por Criticidade e Curva")

# Matriz default: 95% para todas
matriz_default = pd.DataFrame(
    data=95,
    index=['X', 'Y', 'Z'],
    columns=['A', 'B', 'C']
)

# Editor de dados para a matriz
matriz_editada = st.data_editor(
    matriz_default,
    column_config={
        "A": st.column_config.NumberColumn(
            "Curva A", min_value=50, max_value=99, step=1
        ),
        "B": st.column_config.NumberColumn(
            "Curva B", min_value=50, max_value=99, step=1
        ),
        "C": st.column_config.NumberColumn(
            "Curva C", min_value=50, max_value=99, step=1
        ),
    },
    disabled=False,
    hide_index=False,
    num_rows="fixed"
)

st.info("Edite os valores acima (50% a 99%). Linhas: Criticidade (X,Y,Z). Colunas: Curva (A,B,C).")

# ==============================
# FILTROS
# ==============================
st.header("Filtros")
col_emp, col_busca = st.columns(2)
with col_emp:
    empresas = st.multiselect("Empresas", options=sorted(df['Empresa'].unique()), default=sorted(df['Empresa'].unique()))

with col_busca:
    busca = st.text_input("Buscar SKU ou descrição")

df_filtrado = df[df['Empresa'].isin(empresas)].copy()

if busca:
    mask = (df_filtrado['SKU'].str.contains(busca, case=False, na=False) |
            df_filtrado['Descrição_do_Material'].str.contains(busca, case=False, na=False))
    df_filtrado = df_filtrado[mask]

skus_selecionados = st.multiselect("SKUs específicos (vazio = todos)", options=sorted(df_filtrado['SKU'].unique()))

# ==============================
# CÁLCULO
# ==============================
if st.button("Calcular Estoque de Segurança", type="primary", use_container_width=True):

    resultado = df_filtrado.copy()
    if skus_selecionados:
        resultado = resultado[resultado['SKU'].isin(skus_selecionados)]

    if resultado.empty:
        st.warning("Nenhum SKU encontrado com os filtros.")
        st.stop()

    # Cálculo individual por SKU, baseado na matriz
    resultado['Nível de Serviço (%)'] = 0
    resultado['z'] = 0.0
    resultado['SS_Calculado'] = 0.0
    resultado['SS_Arredondado'] = 0

    for idx, row in resultado.iterrows():
        crit = row['Criticidade']
        curva = row['Curva']

        if crit in matriz_editada.index and curva in matriz_editada.columns:
            nivel = matriz_editada.loc[crit, curva]
            z_val = norm.ppf(nivel / 100.0)
        else:
            nivel = 95  # Default se não mapeado
            z_val = norm.ppf(0.95)
            st.warning(f"SKU {row['SKU']}: Criticidade '{crit}' ou Curva '{curva}' não mapeada. Usando default 95%.")

        resultado.at[idx, 'Nível de Serviço (%)'] = nivel
        resultado.at[idx, 'z'] = z_val

        periodo = row['Lead Time médio (meses)'] + 1
        var_dem = periodo * (row['Desvio Padrão Consumo (un)'] ** 2)
        var_lt = (row['Consumo Médio Mensal (un)'] ** 2) * (row['Desvio Padrão LT (meses)'] ** 2)

        ss_calc = z_val * np.sqrt(var_dem + var_lt)
        resultado.at[idx, 'SS_Calculado'] = ss_calc
        resultado.at[idx, 'SS_Arredondado'] = np.ceil(ss_calc).astype(int)

    # Cobertura em meses
    consumo = resultado['Consumo Médio Mensal (un)']
    resultado['Cobertura_Atual_Meses'] = np.where(consumo > 0, resultado['Estoque de Segurança (unidades)'] / consumo, 0)
    resultado['Cobertura_Calculada_Meses'] = np.where(consumo > 0, resultado['SS_Arredondado'] / consumo, 0)

    # Valor
    resultado['Valor_SS_Calculado'] = resultado['SS_Arredondado'] * resultado['Valor Unit']

    # Diferenças
    resultado['Diferença_Unidades'] = resultado['SS_Arredondado'] - resultado['Estoque de Segurança (unidades)']
    resultado['Diferença_%'] = np.where(
        resultado['Estoque de Segurança (unidades)'] > 0,
        ((resultado['SS_Arredondado'] / resultado['Estoque de Segurança (unidades)']) - 1) * 100,
        np.nan
    )

    # Tabela final (incluindo novas colunas)
    tabela = resultado[[
        'Empresa', 'Classe', 'SKU', 'Descrição_do_Material',
        'Criticidade', 'Curva', 'Nível de Serviço (%)',
        'Lead Time médio (meses)', 'Consumo Médio Mensal (un)',
        'Estoque de Segurança (unidades)', 'SS_Arredondado',
        'Diferença_Unidades', 'Diferença_%',
        'Cobertura_Atual_Meses', 'Cobertura_Calculada_Meses',
        'Valor Unit', 'Valor_SS_Calculado'
    ]].round(2)

    tabela = tabela.rename(columns={
        'Estoque de Segurança (unidades)': 'SS Atual',
        'SS_Arredondado': 'SS Calculado',
        'Valor_SS_Calculado': 'Valor SS Calculado (R$)',
        'Cobertura_Atual_Meses': 'Cobertura Atual (meses)',
        'Cobertura_Calculada_Meses': 'Cobertura Calculada (meses)'
    })

    st.markdown(f"### Resultados — {len(tabela)} SKUs")
    st.dataframe(tabela.style.format({
        'Valor SS Calculado (R$)': 'R$ {:,.0f}',
        'Diferença_%': '{:+.1f}%',
        'Cobertura Atual (meses)': '{:.2f}',
        'Cobertura Calculada (meses)': '{:.2f}',
        'Valor Unit': 'R$ {:.2f}',
        'Nível de Serviço (%)': '{:.0f}%'
    }), use_container_width=True)

    # Resumo financeiro
    c1, c2, c3 = st.columns(3)
    valor_atual = (tabela['SS Atual'] * tabela['Valor Unit']).sum()
    valor_novo = tabela['Valor SS Calculado (R$)'].sum()
    impacto = valor_novo - valor_atual

    c1.metric("Valor atual do SS", f"R$ {valor_atual:,.0f}")
    c2.metric("Valor calculado do SS", f"R$ {valor_novo:,.0f}")
    c3.metric("Impacto financeiro", f"R$ {impacto:,.0f}",
              delta=f"{impacto/valor_atual*100 if valor_atual>0 else 0:+.1f}%")

    # ========================
    # MATRIZES DE RESUMO
    # ========================
    st.markdown("---")
    st.subheader("Resumo por Criticidade (X,Y,Z) e Curva (A,B,C)")

    # Matriz 1: Valor Total do Estoque de Segurança Calculado
    matriz_valor = pd.pivot_table(
        tabela,
        values='Valor SS Calculado (R$)',
        index='Criticidade',
        columns='Curva',
        aggfunc='sum',
        fill_value=0,
        margins=True,
        margins_name='Total'
    )
    matriz_valor = matriz_valor.reindex(['X', 'Y', 'Z', 'Total'])
    matriz_valor.columns = pd.Index(list(matriz_valor.columns[:-1]) + ['Total']) if 'Total' in matriz_valor.columns else matriz_valor.columns
    st.write("**Matriz: Valor Total do Estoque de Segurança Calculado (R$)**")
    st.dataframe(matriz_valor.style.format("R$ {:,.0f}"), use_container_width=True)

    # Matriz 2: Cobertura Total Calculada (soma dos meses)
    matriz_cobertura = pd.pivot_table(
        tabela,
        values='Cobertura Calculada (meses)',
        index='Criticidade',
        columns='Curva',
        aggfunc='mean',
        fill_value=0,
        margins=True,
        margins_name='Total'
    )
    matriz_cobertura = matriz_cobertura.reindex(['X', 'Y', 'Z', 'Total'])
    matriz_cobertura.columns = pd.Index(list(matriz_cobertura.columns[:-1]) + ['Total']) if 'Total' in matriz_cobertura.columns else matriz_cobertura.columns
    st.write("**Matriz: Cobertura Total Calculada **")
    st.dataframe(matriz_cobertura.style.format("{:.2f}"), use_container_width=True)

    # Matriz 3: Quantidade de SKUs
    matriz_count = pd.pivot_table(
        tabela,
        values='SKU',
        index='Criticidade',
        columns='Curva',
        aggfunc='count',
        fill_value=0,
        margins=True,
        margins_name='Total'
    )
    matriz_count = matriz_count.reindex(['X', 'Y', 'Z', 'Total'])
    matriz_count.columns = pd.Index(list(matriz_count.columns[:-1]) + ['Total']) if 'Total' in matriz_count.columns else matriz_count.columns
    st.write("**Matriz: Quantidade de SKUs por Mix**")
    st.dataframe(matriz_count, use_container_width=True)

    # ========================
    # DOWNLOAD DO RESULTADO
    # ========================
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        tabela.to_excel(writer, index=False, sheet_name='Resultado')
    buffer.seek(0)

    st.download_button(
        label="Baixar resultado completo em Excel",
        data=buffer.getvalue(),
        file_name="Estoque_Seguranca_Resultado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.success("Pronto! Configure a matriz, ajuste os filtros e calcule.")
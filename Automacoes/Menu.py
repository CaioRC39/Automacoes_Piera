# ==============================================================================
# HUB DE FERRAMENTAS DE AUTOMAÇÃO
# ==============================================================================
# Este é o arquivo principal (página de boas-vindas) do seu aplicativo
# de múltiplas páginas no Streamlit.
# ==============================================================================

import streamlit as st

st.set_page_config(
    page_title="Hub de Ferramentas de Automação",
    page_icon="🤖",
    layout="wide"
)

st.title("Bem-vindo ao Hub de Ferramentas de Automação 👋")
st.sidebar.success("Selecione uma ferramenta acima.")

st.markdown(
    """
    Este é um hub central para as ferramentas de automação de projetos de PD&I.

    **👈 Selecione uma ferramenta na barra lateral** para começar!

    ### Ferramentas Disponíveis:
    - **Preenchedor de Planilha:** Preenche a planilha base a partir de múltiplos arquivos.
    - **Gerador de Relatório:** Cria um relatório consolidado (LP, RH e ST) a partir da planilha de Valoração e dos TAs.
    """
)

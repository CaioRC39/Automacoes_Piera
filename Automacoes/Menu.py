# ==============================================================================
# MENU DE FERRAMENTAS DE AUTOMAÇÃO
# ==============================================================================
# Este é o arquivo principal (página de boas-vindas) do seu aplicativo
# de múltiplas páginas no Streamlit.
# ==============================================================================

import streamlit as st

st.set_page_config(
    page_title="Menu de Ferramentas de Automação",
    page_icon="🤖",
    layout="wide"
)

st.title("Bem-vindo ao Menu de Ferramentas de Automação da Piera! 👋")
st.sidebar.success("Selecione uma ferramenta acima.")

st.markdown(
    """
    Este é um menu central para as ferramentas de automação de projetos de PD&I.

    **👈 Selecione uma ferramenta na barra lateral** para começar!

    ### Ferramentas Disponíveis:
    - **Preenchedor de NewPiit:** Preenche o NewPiit a partir de múltiplos arquivos.
    - **Gerador de planilha com LP, RH e ST:** Cria uma planilha consolidada com Linhas de Pesquisa, Recursos Humanos e Serviços de Terceiros e Viagens a partir da planilha de Valoração e dos TAs.
    - **Formatador para texto de NewPiit:** Gera um texto formatado de uma aba selecionada do NewPiit, com direito a filtros.
    """
)



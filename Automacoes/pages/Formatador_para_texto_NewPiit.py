# -*- coding: utf-8 -*-

# PASSO 1: Importar as bibliotecas necessárias
import streamlit as st
import pandas as pd
from thefuzz import process
import io


# ==============================================================================
#                    MAPEAMENTO INTELIGENTE DE COLUNAS
# ==============================================================================
# Todas as nossas funções auxiliares de lógica pura são mantidas aqui.

def normalizar_nome_coluna(nome):
    """Função de limpeza para padronizar os nomes de colunas."""
    if not isinstance(nome, str):
        return ''
    return nome.lower().strip()

def mapear_colunas_similares(colunas_da_planilha, colunas_esperadas, limiar=80):
    """(FERRAMENTA INTERNA) Encontra colunas por similaridade (fuzzy match) para typos."""
    mapeamento = {}
    colunas_nao_encontradas = []
    mapa_reais_normalizadas = {normalizar_nome_coluna(c): c for c in colunas_da_planilha}
    colunas_reais_normalizadas = list(mapa_reais_normalizadas.keys())

    for nome_esperado in colunas_esperadas:
        nome_esperado_normalizado = normalizar_nome_coluna(nome_esperado)
        melhor_match_normalizado, pontuacao = process.extractOne(nome_esperado_normalizado, colunas_reais_normalizadas)

        if pontuacao >= limiar:
            nome_real_encontrado = mapa_reais_normalizadas[melhor_match_normalizado]
            mapeamento[nome_esperado] = nome_real_encontrado
        else:
            colunas_nao_encontradas.append(nome_esperado)
    return mapeamento, colunas_nao_encontradas

def mapear_colunas_nativas(colunas_da_planilha, colunas_esperadas):
    """(FERRAMENTA INTERNA) Encontra colunas por correspondência exata ou por 'contém'."""
    mapeamento = {}
    colunas_nao_encontradas = []
    mapa_reais_normalizadas = {normalizar_nome_coluna(c): c for c in colunas_da_planilha}
    colunas_reais_normalizadas = list(mapa_reais_normalizadas.keys())
    colunas_ja_mapeadas = []

    for nome_esperado in colunas_esperadas:
        nome_esperado_normalizado = normalizar_nome_coluna(nome_esperado)
        melhor_match_encontrado = None

        if nome_esperado_normalizado in colunas_reais_normalizadas and nome_esperado_normalizado not in colunas_ja_mapeadas:
            melhor_match_encontrado = nome_esperado_normalizado
        else:
            candidatos = [r for r in colunas_reais_normalizadas if r not in colunas_ja_mapeadas and (nome_esperado_normalizado in r or r in nome_esperado_normalizado)]
            if candidatos:
                melhor_match_encontrado = min(candidatos, key=lambda real: abs(len(real) - len(nome_esperado_normalizado)))

        if melhor_match_encontrado:
            nome_real_original = mapa_reais_normalizadas[melhor_match_encontrado]
            mapeamento[nome_esperado] = nome_real_original
            colunas_ja_mapeadas.append(melhor_match_encontrado)
        else:
            colunas_nao_encontradas.append(nome_esperado)
    return mapeamento, colunas_nao_encontradas

def mapear_colunas_inteligentemente(colunas_da_planilha, colunas_esperadas, limiar_fuzzy=80):
    """
    FUNÇÃO PRINCIPAL DE MAPEAMENTO. Combina os métodos para máxima precisão.
    É a única função que você precisa chamar.
    """

    mapeamento_inicial, nao_encontradas_inicial = mapear_colunas_nativas(colunas_da_planilha, colunas_esperadas)

    if not nao_encontradas_inicial:
        return mapeamento_inicial, []

    colunas_reais_restantes = [c for c in colunas_da_planilha if c not in mapeamento_inicial.values()]
    colunas_esperadas_restantes = nao_encontradas_inicial

    mapeamento_fuzzy, nao_encontradas_final = mapear_colunas_similares(
        colunas_reais_restantes,
        colunas_esperadas_restantes,
        limiar=limiar_fuzzy
    )

    mapeamento_final = {**mapeamento_inicial, **mapeamento_fuzzy}

    return mapeamento_final, nao_encontradas_final

def formatar_cpf(cpf):
    """
    Recebe um CPF como string, formata para XXX.XXX.XXX-XX.
    Adiciona um '0' à esquerda se tiver 10 dígitos.
    Se inválido, retorna o valor original.
    """
    # Passo 1: Limpa o CPF, removendo qualquer caractere que não seja um dígito.
    cpf_limpo = ''.join(filter(str.isdigit, str(cpf)))

    # Passo 2: NOVO - Verifica se o CPF tem 10 dígitos e, se tiver, adiciona um zero à esquerda.
    if len(cpf_limpo) == 10:
        cpf_limpo = '0' + cpf_limpo

    # Passo 3: Agora, verifica se o CPF (potencialmente corrigido) tem 11 dígitos para formatar.
    if len(cpf_limpo) == 11:
        # Aplica a máscara de formatação.
        return f'{cpf_limpo[:3]}.{cpf_limpo[3:6]}.{cpf_limpo[6:9]}-{cpf_limpo[9:]}'
    else:
        # Se o CPF original não tinha 10 ou 11 dígitos, retorna o valor original sem formatação.
        return cpf

# ==============================================================================
#           FUNÇÃO PARA PROCESSAR A ABA "RH" (VERSÃO PARA STREAMLIT)
# ==============================================================================
def processar_aba_rh(df, mapeamento, projeto_selecionado):
    """
    Recebe um DataFrame da aba RH, o mapeamento de colunas e o projeto selecionado,
    e RETORNA uma única string com os dados formatados para exibição.
    """
    # --- 1. LÓGICA DE FILTRO ---
    # Define o nome "ideal" da coluna de projeto para buscar no dicionário de mapeamento
    coluna_projeto_ideal = 'Nome da atividade de PD&I (Nome do projeto igual no GERAL)'
    # Pega o nome "real" da coluna que foi encontrado na planilha
    coluna_projeto_real = mapeamento[coluna_projeto_ideal]

    # Filtra o DataFrame se um projeto específico foi escolhido no menu do Streamlit
    if projeto_selecionado != "Listar TODOS os colaboradores":
        df = df[df[coluna_projeto_real] == projeto_selecionado].copy()

    # --- 2. PREPARAÇÃO FINAL DOS DADOS ---
    # Garante que os valores numéricos vazios sejam tratados como 0 e o resto como texto vazio
    df[mapeamento['Valor (R$)']] = df[mapeamento['Valor (R$)']].fillna(0)
    df[mapeamento['Total Horas (Anual)']] = df[mapeamento['Total Horas (Anual)']].fillna(0)
    df = df.fillna('')

    # Se o DataFrame ficar vazio após o filtro, retorna uma mensagem amigável
    if df.empty:
        return "Nenhum colaborador encontrado para a seleção feita."

    # --- 3. GERAÇÃO DO TEXTO DE SAÍDA ---
    # Cria uma lista vazia para armazenar cada linha do nosso texto final
    texto_saida = []
    contador_colaborador = 1

    # Itera sobre as linhas do DataFrame (que pode estar filtrado ou não)
    for indice, linha in df.iterrows():
        # Pula a linha se o nome do colaborador estiver em branco
        nome = str(linha[mapeamento["NOME"]]).strip()
        if not nome:
            continue

        # Extrai todos os outros dados usando o mapeamento
        projeto   = str(linha[mapeamento[coluna_projeto_ideal]]).strip()
        cpf       = formatar_cpf(linha[mapeamento["CPF"]])
        titulacao = str(linha[mapeamento["TITULAÇÃO"]]).strip()
        funcao    = str(linha[mapeamento["FUNÇÃO"]]).strip()
        sexo      = str(linha[mapeamento["SEXO"]]).strip()
        horas     = linha[mapeamento["Total Horas (Anual)"]]
        dedicacao = str(linha[mapeamento["DEDICAÇÃO"]]).strip()
        valor     = linha[mapeamento["Valor (R$)"]]
        atividade = str(linha[mapeamento['Descreva as atividades realizadas pelo profissional (cargo, atividades exercidas e contribuições no projeto)']]).strip()

        # Formata os valores numéricos
        horas_formatado = f"{horas:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if horas != 0 else ''
        valor_formatado = f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if valor != 0 else ''

        # Adiciona cada linha formatada à nossa lista de saída
        texto_saida.append(f"Colaborador {contador_colaborador}")
        texto_saida.append(f"Projeto: {projeto}")
        texto_saida.append(f"CPF: {cpf}")
        texto_saida.append(f"Nome: {nome}")
        texto_saida.append(f"Titulação: {titulacao}")
        texto_saida.append(f"Função: {funcao}")
        texto_saida.append(f"Sexo: {sexo}")
        texto_saida.append(f"Total Horas (Anual): {horas_formatado}")
        texto_saida.append(f"Dedicação: {dedicacao}")
        texto_saida.append(f"Valor: {valor_formatado}")
        texto_saida.append(f"Atividade: {atividade}")
        texto_saida.append("-" * 30)

        contador_colaborador += 1

    # --- 4. RETORNO DO RESULTADO ---
    # Junta todas as linhas da lista em uma única string, separadas por quebra de linha
    return "\n".join(texto_saida)

# ==============================================================================
#           FUNÇÃO PARA PROCESSAR A ABA "GERAL" (VERSÃO PARA STREAMLIT)
# ==============================================================================
def processar_aba_geral(df, mapeamento, projeto_selecionado):
    """
    Recebe um DataFrame da aba GERAL, o mapeamento e o projeto selecionado,
    e RETORNA uma string com os dados formatados para exibição.
    """
    # --- 1. LÓGICA DE FILTRO ---
    coluna_projeto_ideal = "Nome da atividade de PD&I: "
    coluna_projeto_real = mapeamento[coluna_projeto_ideal]

    # Filtra o DataFrame se um projeto específico foi escolhido
    if projeto_selecionado != "Listar TODOS os projetos":
        df = df[df[coluna_projeto_real] == projeto_selecionado].copy()

    # --- 2. PREPARAÇÃO FINAL DOS DADOS ---
    df = df.fillna('')
    if df.empty:
        return "Nenhum projeto encontrado para a seleção feita."

    # --- 3. GERAÇÃO DO TEXTO DE SAÍDA ---
    texto_saida = []
    contador_projeto = 1

    for indice, linha in df.iterrows():
        projeto = str(linha[mapeamento[coluna_projeto_ideal]]).strip()
        if not projeto or projeto.lower() == 'nan':
            continue

        # Extração de todos os dados da linha usando o mapeamento
        descricao_projeto = str(linha[mapeamento["Descrição do Projeto:"]]).strip()
        pb_pa_de = str(linha[mapeamento["PB, PA ou DE:"]]).strip()
        area_projeto = str(linha[mapeamento["Área do Projeto:"]]).strip()
        palavras_chave = str(linha[mapeamento["Palavras-Chave (Separadas por vírgula):"]]).strip()
        natureza = str(linha[mapeamento["Natureza (Produto, Processo ou Serviço):"]]).strip()
        elemento_novo = str(linha[mapeamento["Destaque o elemento tecnologicamente novo ou inovador da atividade: "]]).strip()
        barreiras = str(linha[mapeamento["Qual a barreira ou desafio tecnológico superável: "]]).strip()
        metodologia = str(linha[mapeamento["Qual a metodologia / métodos utilizados: "]]).strip()
        atividade_continua = str(linha[mapeamento["A atividade é contínua (ciclo de vida maior que 1 ano)?  (Sim ou Não)"]]).strip()
        data_inicio = str(linha[mapeamento["Data de início: (formato dd/mm/aaaa)"]]).strip()
        previsao_termino = str(linha[mapeamento["Previsão de término: (formato dd/mm/aaaa)"]]).strip()
        atividade_ano_base = str(linha[mapeamento["Caso a atividade/projeto seja continuada, informar Atividade de PD&I desenvolvida no ano-base"]]).strip()
        info_complementares = str(linha[mapeamento["Descrição Complementar: "]]).strip()
        resultado_economico = str(linha[mapeamento["Resultado Econômico:"]]).strip()
        resultado_inovacao = str(linha[mapeamento["Resultado de Inovação:"]]).strip()
        trl_inicial = str(linha[mapeamento["TRL Inicial"]]).strip()
        trl_final = str(linha[mapeamento["TRL Final"]]).strip()
        justificativa_trl = str(linha[mapeamento["Justificativa TRL"]]).strip()
        ods = str(linha[mapeamento["ODS"]]).strip()
        justificativa_ods = str(linha[mapeamento["Justificativa ODS"]]).strip()
        alinha_politicas = str(linha[mapeamento["Os projetos de PD&I da empresa se alinham com as políticas públicas nacionais? (Sim ou Não)"]]).strip()
        desc_alinha_politicas = str(linha[mapeamento["Alinhamento do Projeto com Políticas, Programas e Estratégias Governamentais"]]).strip()

        # Adiciona as linhas formatadas à lista de saída, respeitando as condicionais
        texto_saida.append(f"--- Projeto {contador_projeto} ---")
        texto_saida.append(f"Projeto: {projeto}")
        texto_saida.append(f"Descrição do Projeto: {descricao_projeto}")
        texto_saida.append(f"PB, PA ou DE: {pb_pa_de}")
        texto_saida.append(f"Área do Projeto: {area_projeto}")
        texto_saida.append(f"Palavras-Chave: {palavras_chave}")
        texto_saida.append(f"Natureza: {natureza}")
        texto_saida.append(f"Elemento Novo: {elemento_novo}")
        texto_saida.append(f"Barreiras: {barreiras}")
        texto_saida.append(f"Metodologia: {metodologia}")

        texto_saida.append(f"Atividade contínua?: {atividade_continua}")
        if atividade_continua.strip().lower() != 'não':
            texto_saida.append(f"Data de início: {data_inicio}")
            texto_saida.append(f"Previsão de término: {previsao_termino}")
            texto_saida.append(f"Atividade de PD&I desenvolvida no ano-base: {atividade_ano_base}")

        texto_saida.append(f"Informações Complementares: {info_complementares}")
        texto_saida.append(f"Resultado Econômico: {resultado_economico}")
        texto_saida.append(f"Resultado de Inovação: {resultado_inovacao}")
        texto_saida.append(f"TRL Inicial: {trl_inicial}")
        texto_saida.append(f"TRL Final: {trl_final}")
        texto_saida.append(f"Justificativa TRL: {justificativa_trl}")
        texto_saida.append(f"ODS: {ods}")
        texto_saida.append(f"Justificativa ODS: {justificativa_ods}")

        texto_saida.append(f"Alinha-se às políticas públicas?: {alinha_politicas}")
        if alinha_politicas.strip().lower() != 'não':
            texto_saida.append(f"Descrição alinhamento às Políticas Públicas: {desc_alinha_politicas}")

        texto_saida.append("-" * 30)
        contador_projeto += 1

    # --- 4. RETORNO DO RESULTADO ---
    return "\n".join(texto_saida)

# ==============================================================================
#           FUNÇÃO PARA PROCESSAR A ABA "DISPÊNDIOS ST" (VERSÃO PARA STREAMLIT)
# ==============================================================================
def processar_aba_dispêndios_st(df, mapeamento, projeto_selecionado):
    """
    Recebe um DataFrame da aba DISPÊNDIOS ST, o mapeamento e o projeto selecionado,
    e RETORNA uma string com os dados formatados para exibição.
    """
    # --- 1. LÓGICA DE FILTRO ---
    coluna_projeto_ideal = 'Nome da atividade de PD&I (Nome do projeto igual no GERAL)'
    coluna_projeto_real = mapeamento[coluna_projeto_ideal]

    # Filtra o DataFrame se um projeto específico foi escolhido
    if projeto_selecionado != "Listar TODOS os dispêndios":
        df = df[df[coluna_projeto_real] == projeto_selecionado].copy()

    # --- 2. PREPARAÇÃO FINAL DOS DADOS ---
    df[mapeamento['Valor Total']] = df[mapeamento['Valor Total']].fillna(0)
    df = df.fillna('')

    if df.empty:
        return "Nenhum dispêndio de Serviço de Terceiro e Viagens encontrado para a seleção feita."

    # --- 3. GERAÇÃO DO TEXTO DE SAÍDA ---
    texto_saida = []
    contador_dispêndio = 1

    for indice, linha in df.iterrows():
        razao_social = str(linha[mapeamento["Prestador de Serviço"]]).strip()
        if not razao_social or razao_social.lower() == 'nan':
            continue

        # Extração de todos os dados da linha usando o mapeamento
        projeto = str(linha[mapeamento[coluna_projeto_ideal]]).strip()
        porte_tipo = str(linha[mapeamento["TIPO"]]).strip()
        situacao = str(linha[mapeamento["Situação (Contratado, Em Execução, Terminado)"]]).strip()
        cnpj_cpf = str(linha[mapeamento["CNPJ/CPF"]]).strip()
        caracterizacao = str(linha[mapeamento["Caracterizar o Serviço Realizado"]]).strip()
        valor_total = linha[mapeamento["Valor Total"]]
        centro_pesquisa = str(linha[mapeamento["Centro, departamento ou grupo de pesquisa da universidade/instituição de pesquisa contratada "]]).strip()
        centro_embrapii = str(linha[mapeamento["Centro, Departamento ou Grupo de Pesquisa (caso seja credenciada Embrapii)"]]).strip()
        codigo_embrapii = str(linha[mapeamento["Código do projeto Embrapii (caso seja credenciada Embrapii)"]]).strip()

        # Formatação do valor monetário
        valor_formatado = f"R$ {valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if valor_total != 0 else ''

        # Adiciona as linhas formatadas à lista de saída
        texto_saida.append(f"Dispêndio {contador_dispêndio}")
        texto_saida.append(f"Projeto: {projeto}")
        texto_saida.append(f"Porte/Tipo de serviço: {porte_tipo}")
        texto_saida.append(f"Situação: {situacao}")
        texto_saida.append(f"Razão Social: {razao_social}")
        texto_saida.append(f"CNPJ/CPF: {cnpj_cpf}")
        texto_saida.append(f"Caracterização do Serviço Realizado: {caracterizacao}")
        texto_saida.append(f"Valor Total: {valor_formatado}")
        texto_saida.append(f"Centro, departamento ou grupo de pesquisa da universidade/instituição de pesquisa contratada: {centro_pesquisa}")
        texto_saida.append(f"Centro, Departamento ou Grupo de Pesquisa (caso seja credenciada Embrapii): {centro_embrapii}")
        texto_saida.append(f"Código do projeto Embrapii (caso seja credenciada Embrapii): {codigo_embrapii}")
        texto_saida.append("-" * 30)

        contador_dispêndio += 1

    # --- 4. RETORNO DO RESULTADO ---
    return "\n".join(texto_saida)

# ==============================================================================
#           FUNÇÃO PARA PROCESSAR A ABA "DISPÊNDIOS MC" (VERSÃO PARA STREAMLIT)
# ==============================================================================
def processar_aba_dispêndios_mc(df, mapeamento, projeto_selecionado):
    """
    Recebe um DataFrame da aba DISPÊNDIOS MC, o mapeamento e o projeto selecionado,
    e RETORNA uma string com os dados formatados para exibição.
    """
    # --- 1. LÓGICA DE FILTRO ---
    coluna_projeto_ideal = 'Nome da atividade de PD&I (Nome do projeto igual no GERAL)'
    coluna_projeto_real = mapeamento[coluna_projeto_ideal]

    # Filtra o DataFrame se um projeto específico foi escolhido
    if projeto_selecionado != "Listar TODOS os dispêndios de materiais":
        df = df[df[coluna_projeto_real] == projeto_selecionado].copy()

    # --- 2. PREPARAÇÃO FINAL DOS DADOS ---
    df[mapeamento['Valor Total']] = df[mapeamento['Valor Total']].fillna(0)
    df = df.fillna('')

    if df.empty:
        return "Nenhum dispêndio de Material de Cosumo encontrado para a seleção feita."

    # --- 3. GERAÇÃO DO TEXTO DE SAÍDA ---
    texto_saida = []
    contador_dispêndio = 1

    for indice, linha in df.iterrows():
        material = str(linha[mapeamento["Identificação do Material"]]).strip()
        if not material or material.lower() == 'nan':
            continue

        # Extração dos dados da linha usando o mapeamento
        projeto = str(linha[mapeamento[coluna_projeto_ideal]]).strip()
        descricao = str(linha[mapeamento["Descrição"]]).strip()
        valor_total = linha[mapeamento["Valor Total"]]

        # Formatação do valor monetário
        valor_formatado = f"R$ {valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if valor_total != 0 else ''

        # Adiciona as linhas formatadas à lista de saída
        texto_saida.append(f"Dispêndio MC {contador_dispêndio}")
        texto_saida.append(f"Projeto: {projeto}")
        texto_saida.append(f"Material: {material}")
        texto_saida.append(f"Descrição: {descricao}")
        texto_saida.append(f"Valor Total: {valor_formatado}")
        texto_saida.append("-" * 30)

        contador_dispêndio += 1

    # --- 4. RETORNO DO RESULTADO ---
    return "\n".join(texto_saida)

# ==============================================================================
#           APLICAÇÃO STREAMLIT (INTERFACE GRÁFICA PRINCIPAL)
# ==============================================================================

st.set_page_config(page_title="Formatador de Relatórios", layout="wide")
st.title("📄 Formatador para texto de NewPiit")

texto_instrucoes = """
### ✨ Bem-vindo ao Formatador para texto de NewPiit!

Esta ferramenta foi projetada para simplificar sua vida! Ela automatiza a tediosa tarefa de copiar e colar informações do NewPiit, transformando os dados brutos em textos formatados e prontos para serem usados.

---

### 🚀 Como Começar (Passo a Passo)

**1. Carregue o NewPiit _preenchido_**
   - Arraste e solte o NewPiit _já preenchido_ na área indicada ou clique no botão **"Browse files"** para procurá-lo no seu computador.
   - O aplicativo aceita apenas arquivos no formato `.xlsx`.

**2. Selecione a Aba**
   - No primeiro menu suspenso que aparecer, escolha qual aba do NewPiit você deseja processar (ex: "Recursos Humanos (aba RH)", "Informações dos projetos (aba GERAL)", etc.).

**3. Filtre por Projeto (Opcional)**
   - Se a aba selecionada contiver informações de múltiplos projetos, um segundo menu aparecerá.
   - Você pode escolher um projeto específico para ver apenas os dados relacionados a ele, ou selecionar a primeira opção (**"Listar TODOS..."**) para processar os dados de todos os projetos daquela aba.

**4. Gere o Texto Formatado**
   - Após fazer suas seleções, clique no botão azul principal (ex: **"✨ Gerar Texto da Aba '[Nome da aba selecionada]'"**).
   - Aguarde alguns instantes enquanto a mágica acontece!

**5. Copie o Resultado**
   - O texto final, perfeitamente formatado, aparecerá em uma grande caixa de texto na parte inferior da página.
   - Basta clicar no botão de copiar no canto superior direito da caixa e colar onde você precisar!
   - Ou você pode também copiar todo o conteúdo com `Ctrl+C` no Windows ou `Cmd+C` Mac e colar onde você precisar!
"""

st.markdown(texto_instrucoes)

st.info("**Instruções:**\n1. Faça o upload do NewPiit.\n2. Selecione a aba que deseja processar.\n3. Se aplicável, filtre por um projeto específico.\n4. Clique no botão para gerar o texto formatado.")

# --- PAINEL DE CONTROLE CENTRAL ---
# Este dicionário é o cérebro do app. Ele diz ao Streamlit tudo o que ele precisa
# saber sobre cada aba: qual o nome da planilha, quais colunas esperar, qual
# função de processamento chamar, etc.
CONFIG_ABAS = {
    "Informações dos projetos (Aba GERAL)": {
        "sheet_name": "GERAL",
        "skiprows": 9,
        "funcao_processamento": processar_aba_geral,
        "filtro_projeto": True,
        "coluna_filtro_ideal": "Nome da atividade de PD&I: ",
        "label_filtro_todos": "Listar TODOS os projetos",
        "colunas_esperadas": [
            "Nome da atividade de PD&I: ", "Descrição do Projeto:", "PB, PA ou DE:", "Área do Projeto:",
            "Palavras-Chave (Separadas por vírgula):", "Natureza (Produto, Processo ou Serviço):",
            "Destaque o elemento tecnologicamente novo ou inovador da atividade: ",
            "Qual a barreira ou desafio tecnológico superável: ", "Qual a metodologia / métodos utilizados: ",
            "A atividade é contínua (ciclo de vida maior que 1 ano)?  (Sim ou Não)",
            "Data de início: (formato dd/mm/aaaa)", "Previsão de término: (formato dd/mm/aaaa)",
            "Caso a atividade/projeto seja continuada, informar Atividade de PD&I desenvolvida no ano-base",
            "Descrição Complementar: ", "Resultado Econômico:", "Resultado de Inovação:", "TRL Inicial", "TRL Final",
            "Justificativa TRL", "ODS", "Justificativa ODS",
            "Os projetos de PD&I da empresa se alinham com as políticas públicas nacionais? (Sim ou Não)",
            "Alinhamento do Projeto com Políticas, Programas e Estratégias Governamentais"
        ]
    },
    "Serviços de Terceiros e Viagens (Aba DISPÊNDIOS ST)": {
        "sheet_name": "DISPÊNDIOS ST",
        "skiprows": 9,
        "funcao_processamento": processar_aba_dispêndios_st,
        "filtro_projeto": True,
        "coluna_filtro_ideal": "Nome da atividade de PD&I (Nome do projeto igual no GERAL)",
        "label_filtro_todos": "Listar TODOS os dispêndios",
        "colunas_esperadas": [
            'Nome da atividade de PD&I (Nome do projeto igual no GERAL)', 'TIPO', 'Situação (Contratado, Em Execução, Terminado)',
            'Prestador de Serviço', 'CNPJ/CPF', 'Caracterizar o Serviço Realizado', 'Valor Total',
            'Centro, departamento ou grupo de pesquisa da universidade/instituição de pesquisa contratada ',
            'Centro, Departamento ou Grupo de Pesquisa (caso seja credenciada Embrapii)',
            'Código do projeto Embrapii (caso seja credenciada Embrapii)'
        ]
    },
    "Dispêndios com Material de Consumo (Aba DISPÊNDIOS MC)": {
        "sheet_name": "DISPÊNDIOS MC",
        "skiprows": 9,
        "funcao_processamento": processar_aba_dispêndios_mc,
        "filtro_projeto": True,
        "coluna_filtro_ideal": 'Nome da atividade de PD&I (Nome do projeto igual no GERAL)',
        "label_filtro_todos": "Listar TODOS os dispêndios de materiais",
        "colunas_esperadas": [
            'Nome da atividade de PD&I (Nome do projeto igual no GERAL)', 'Identificação do Material', 'Descrição', 'Valor Total'
        ]
    },
    "Informações dos colaboradores (Aba RH)": {
        "sheet_name": "RH",
        "skiprows": 9,
        "funcao_processamento": processar_aba_rh,
        "filtro_projeto": True,
        "coluna_filtro_ideal": 'Nome da atividade de PD&I (Nome do projeto igual no GERAL)',
        "label_filtro_todos": "Listar TODOS os colaboradores",
        "colunas_esperadas": [
            'Nome da atividade de PD&I (Nome do projeto igual no GERAL)', 'CPF', 'NOME', 'TITULAÇÃO', 'FUNÇÃO',
            'SEXO', 'Total Horas (Anual)', 'DEDICAÇÃO', 'Valor (R$)',
            'Descreva as atividades realizadas pelo profissional (cargo, atividades exercidas e contribuições no projeto)'
        ]
    }
}

# --- LÓGICA DA INTERFACE ---
uploaded_file = st.file_uploader("1. Faça o upload do NewPiit (.xlsx)", type="xlsx")

if uploaded_file is not None:
    st.success(f"Arquivo '{uploaded_file.name}' carregado com sucesso!")

    opcoes_abas = list(CONFIG_ABAS.keys())
    aba_selecionada_nome = st.selectbox("2. Selecione a aba que deseja processar:", options=opcoes_abas, index=None, placeholder="Escolha uma opção...")

    if aba_selecionada_nome:
        config = CONFIG_ABAS[aba_selecionada_nome]

        try:
            df = pd.read_excel(uploaded_file, sheet_name=config["sheet_name"], skiprows=config["skiprows"])
            mapeamento, nao_encontradas = mapear_colunas_inteligentemente(df.columns, config["colunas_esperadas"])

            if nao_encontradas:
                st.error(f"**Colunas não encontradas!**\n\nAs seguintes colunas essenciais não foram encontradas na aba '{config['sheet_name']}':\n- {'\n- '.join(nao_encontradas)}\n\nPor favor, verifique sua planilha e tente novamente.")
            else:
                # Pré-processamento de tipos de dados (centralizado aqui)
                if aba_selecionada_nome == "Informações dos projetos (GERAL)":
                    colunas_data = ["Data de início: (formato dd/mm/aaaa)", "Previsão de término: (formato dd/mm/aaaa)"]
                    for col in colunas_data:
                        df[mapeamento[col]] = pd.to_datetime(df[mapeamento[col]], errors='coerce').dt.strftime('%d/%m/%Y')
                elif aba_selecionada_nome == "Informações dos colaboradores (RH)":
                    df[mapeamento['Valor (R$)']] = pd.to_numeric(df[mapeamento['Valor (R$)']], errors='coerce')
                    df[mapeamento['Total Horas (Anual)']] = pd.to_numeric(df[mapeamento['Total Horas (Anual)']], errors='coerce')
                elif "DISPÊNDIOS" in aba_selecionada_nome:
                    df[mapeamento['Valor Total']] = pd.to_numeric(df[mapeamento['Valor Total']], errors='coerce')

                # Menu de filtro de projetos (se aplicável)
                projeto_selecionado = "TODOS" # Valor padrão
                if config.get("filtro_projeto", False):
                    coluna_projeto_real = mapeamento[config["coluna_filtro_ideal"]]
                    lista_projetos = [config["label_filtro_todos"]] + sorted(list(df[coluna_projeto_real].dropna().unique()))
                    projeto_selecionado = st.selectbox("3. (Opcional) Filtre por um projeto:", options=lista_projetos)

                # Botão para iniciar o processamento
                if st.button(f"✨ Gerar Texto da Aba '{aba_selecionada_nome}'", type="primary"):
                    with st.spinner("Processando... Por favor, aguarde."):
                        funcao = config["funcao_processamento"]
                        resultado_texto = funcao(df, mapeamento, projeto_selecionado)
                    
                        st.subheader("Resultado Formatado:")

                        st.code(resultado_texto, language=None, line_numbers=True)
                    
                        st.success("Processamento concluído com sucesso!")

        except Exception as e:
            st.error(f"**Ocorreu um erro ao processar o NewPiit!**\n\nVerifique se a aba '{config['sheet_name']}' existe no seu arquivo e se o formato está correto.\n\nDetalhe do erro: {e}")

else:
    st.warning("Aguardando o upload do NewPiit...")

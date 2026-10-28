"""
Módulo responsável pela leitura inteligente e dinâmica da tabela de entrada.
1. Identifica automaticamente em qual linha e coluna estão os cabeçalhos da planilha
   (sejam os quadrantes SWOT ou quaisquer novas colunas de outras abas/páginas).
2. Ignora títulos mesclados superiores ('MATRIZ SWOT', 'Ambiente Interno', etc.)
   e colunas laterais de identificação ('UNIDADE', 'SESP', 'ÓRGÃO').
"""

import csv
import io
import re
import unicodedata
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import pandas as pd

from src.config import ARQUIVO_ENTRADA_PADRAO, ARQUIVO_LOG_PROGRESSO, DADOS_DIR
from src.driver_windows import ler_do_clipboard

SINONIMOS_SWOT = {
    "FORÇAS": {"FORCAS", "FORCA", "FORÇAS", "FORÇA"},
    "FRAQUEZAS": {"FRAQUEZAS", "FRAQUEZA"},
    "OPORTUNIDADES": {"OPORTUNIDADES", "OPORTUNIDADE"},
    "AMEAÇAS": {"AMEACAS", "AMEACA", "AMEAÇAS", "AMEAÇA"},
    "DESCRIÇÃO": {
        "DESCRICAO",
        "DESCRIÇÃO",
        "DESCRICAO SUSCINTA DE NECESSIDADE PRIORIZADA NO CICLO DO PDTIC",
        "DESCRIÇÃO SUSCINTA DE NECESSIDADE PRIORIZADA NO CICLO DO PDTIC",
        "DESCRICAO SUCINTA DE NECESSIDADE PRIORIZADA NO CICLO DO PDTIC",
        "DESCRIÇÃO SUCINTA DE NECESSIDADE PRIORIZADA NO CICLO DO PDTIC",
        "DESCRICAO DA NECESSIDADE",
        "DESCRIÇÃO DA NECESSIDADE",
    },
    "PRIORIDADE": {
        "PRIORIDADE",
        "CLASSIFICACAO DE PRIORIDADE",
        "CLASSIFICAÇÃO DE PRIORIDADE",
    },
    "ÁREA DEMANDANTE": {
        "AREA DEMANDANTE",
        "ÁREA DEMANDANTE",
        "UNIDADE DEMANDANTE",
        "SETOR DEMANDANTE",
    },
    "JUSTIFICATIVA": {
        "JUSTIFICATIVA",
        "JUSTIFICATIVA DA NECESSIDADE",
    },
    "ID": {
        "ID",
        "ID OE.TIC",
        "ID OE TIC",
        "ID OETIC",
        "OE.TIC",
    },
    "OBJETIVO ESTRATÉGICO": {
        "OBJETIVO ESTRATEGICO",
        "OBJETIVO ESTRATÉGICO",
        "OBJETIVO ESTRATEGICO DE TIC",
        "OBJETIVO ESTRATÉGICO DE TIC",
        "OBJETIVOS ESTRATEGICOS",
        "OBJETIVOS ESTRATÉGICOS",
        "OBJETIVOS ESTRATEGICOS DE TIC",
        "OBJETIVOS ESTRATÉGICOS DE TIC",
    },
    "META": {
        "META",
        "METAS",
        "META DE TIC",
    },
    "AÇÃO PRINCIPAL": {
        "ACAO PRINCIPAL",
        "AÇÃO PRINCIPAL",
        "RESULTADO-CHAVE (KR) ESPERADO (ACAO PRINCIPAL)",
        "RESULTADO-CHAVE (KR) ESPERADO (AÇÃO PRINCIPAL)",
        "RESULTADO CHAVE (KR) ESPERADO (ACAO PRINCIPAL)",
        "RESULTADO CHAVE (KR) ESPERADO (AÇÃO PRINCIPAL)",
        "RESULTADO-CHAVE (KR) ESPERADO",
        "RESULTADO CHAVE (KR) ESPERADO",
        "RESULTADO-CHAVE",
        "RESULTADO CHAVE",
    },
    "INDICADOR": {
        "INDICADOR",
        "INDICADORES",
        "INDICADOR-CHAVE DE DESEMPENHO (KPI)",
        "INDICADOR CHAVE DE DESEMPENHO (KPI)",
        "INDICADOR-CHAVE DE DESEMPENHO",
        "INDICADOR CHAVE DE DESEMPENHO",
        "INDICADORES (KPIS)",
        "KPI",
    },
    "PRAZO": {
        "PRAZO",
        "PRAZOS",
        "PRAZO DE EXECUCAO",
        "PRAZO DE EXECUÇÃO",
        "QUADRANTE",
    },
    "RESPONSÁVEL": {
        "RESPONSAVEL",
        "RESPONSÁVEL",
        "RESPONSAVEIS",
        "RESPONSÁVEIS",
        "AREA RESPONSAVEL",
        "ÁREA RESPONSÁVEL",
    },
    "FÓRMULA / DESCRIÇÃO": {
        "FORMULA / DESCRICAO",
        "FÓRMULA / DESCRIÇÃO",
        "METRICA / FORMULA",
        "MÉTRICA / FÓRMULA",
        "METRICA/FORMULA",
        "MÉTRICA/FÓRMULA",
        "FORMULA/DESCRICAO",
        "FÓRMULA/DESCRIÇÃO",
        "METRICA",
        "MÉTRICA",
        "FORMULA",
        "FÓRMULA",
        "METRICA DE CALCULO",
        "MÉTRICA DE CÁLCULO",
        "FORMULA DE CALCULO",
        "FÓRMULA DE CÁLCULO",
    },
    "AÇÃO (ID)": {
        "ACAO (ID)",
        "AÇÃO (ID)",
        "ID DA ACAO",
        "ID DA AÇÃO",
        "ID ACAO",
        "ID AÇÃO",
        "ACAO ID",
        "AÇÃO ID",
        "CODIGO DA ACAO",
        "CÓDIGO DA AÇÃO",
    },
    "DESCRIÇÃO DO CUSTO": {
        "DESCRICAO DO CUSTO",
        "DESCRIÇÃO DO CUSTO",
        "DESCRICAO CUSTO",
        "DESCRIÇÃO CUSTO",
        "ITEM DE CUSTO",
        "DESCRICAO DA DESPESA",
        "DESCRIÇÃO DA DESPESA",
        "DESPESA",
    },
    "FONTE DE RECURSO": {
        "FONTE DE RECURSO",
        "FONTE DE RECURSOS",
        "FONTE ORCAMENTARIA",
        "FONTE ORÇAMENTÁRIA",
        "FONTE",
    },
    "ANO 1": {"ANO 1", "ANO1", "1º ANO", "1° ANO", "VALOR ANO 1"},
    "ANO 2": {"ANO 2", "ANO2", "2º ANO", "2° ANO", "VALOR ANO 2"},
    "ANO 3": {"ANO 3", "ANO3", "3º ANO", "3° ANO", "VALOR ANO 3"},
    "ANO 4": {"ANO 4", "ANO4", "4º ANO", "4° ANO", "VALOR ANO 4"},
    "RISCO IDENTIFICADO": {
        "RISCO IDENTIFICADO",
        "RISCOS IDENTIFICADOS",
        "RISCO",
        "RISCOS",
        "DESCRICAO DO RISCO",
        "DESCRIÇÃO DO RISCO",
        "EVENTO DE RISCO",
    },
    "PROBABILIDADE": {
        "PROBABILIDADE",
        "PROB",
        "PROBABILIDADE DO RISCO",
    },
    "IMPACTO": {
        "IMPACTO",
        "IMPACTO DO RISCO",
    },
    "NÍVEL": {
        "NIVEL",
        "NÍVEL",
        "NIVEL DE RISCO",
        "NÍVEL DE RISCO",
        "NIVEL DO RISCO",
        "NÍVEL DO RISCO",
        "GRAU DE RISCO",
        "SEVERIDADE",
    },
    "AÇÃO DE MITIGAÇÃO": {
        "ACAO DE MITIGACAO",
        "AÇÃO DE MITIGAÇÃO",
        "ACOES DE MITIGACAO",
        "AÇÕES DE MITIGAÇÃO",
        "MITIGACAO",
        "MITIGAÇÃO",
        "PLANO DE MITIGACAO",
        "PLANO DE MITIGAÇÃO",
        "PLANO DE CONTINGENCIA",
        "PLANO DE CONTINGÊNCIA",
        "RESPOSTA AO RISCO",
    },
    "RESPONSÁVEL PELO RISCO": {
        "RESPONSAVEL PELO RISCO",
        "RESPONSÁVEL PELO RISCO",
        "RESPONSAVEL MITIGACAO",
        "RESPONSÁVEL MITIGAÇÃO",
    },
    "TREINAMENTOS TÉCNICOS": {
        "TREINAMENTOS TECNICOS",
        "TREINAMENTOS TÉCNICOS",
        "TREINAMENTO TECNICO",
        "TREINAMENTO TÉCNICO",
        "CAPACITACAO TECNICA",
        "CAPACITAÇÃO TÉCNICA",
    },
    "AÇÕES DE RETENÇÃO DE TALENTOS": {
        "ACOES DE RETENCAO DE TALENTOS",
        "AÇÕES DE RETENÇÃO DE TALENTOS",
        "ACAO DE RETENCAO DE TALENTOS",
        "AÇÃO DE RETENÇÃO DE TALENTOS",
        "RETENCAO DE TALENTOS",
        "RETENÇÃO DE TALENTOS",
    },
    "AÇÕES MACRO": {
        "ACOES MACRO",
        "AÇÕES MACRO",
        "ACAO MACRO",
        "AÇÃO MACRO",
        "MACRO ACOES",
        "MACRO AÇÕES",
    },
    "AÇÕES MICRO": {
        "ACOES MICRO",
        "AÇÕES MICRO",
        "ACAO MICRO",
        "AÇÃO MICRO",
        "MICRO ACOES",
        "MICRO AÇÕES",
    },
    "AÇÕES EXTRAS": {
        "ACOES EXTRAS",
        "AÇÕES EXTRAS",
        "ACAO EXTRA",
        "AÇÃO EXTRA",
        "EXTRAS",
        "ACOES ADICIONAIS",
        "AÇÕES ADICIONAIS",
    },
}

TITULOS_IGNORADOS = {
    "MATRIZ SWOT",
    "AMBIENTE INTERNO",
    "AMBIENTE EXTERNO",
    "ANALISE SWOT",
    "ANÁLISE SWOT",
    "DIAGNOSTICO",
    "DIAGNÓSTICO",
    "INVENTARIO DE NECESSIDADES",
    "INVENTÁRIO DE NECESSIDADES",
    "3.2. INVENTARIO DE NECESSIDADES",
    "3.2. INVENTÁRIO DE NECESSIDADES",
    "METAS E ACOES",
    "METAS E AÇÕES",
    "4.1. METAS E ACOES",
    "4.1. METAS E AÇÕES",
    "PLANO ORCAMENTARIO",
    "PLANO ORÇAMENTÁRIO",
    "PLANO ORCAMENTARIO E DE INVESTIMENTOS",
    "PLANO ORÇAMENTÁRIO E DE INVESTIMENTOS",
    "4.2. PLANO ORCAMENTARIO",
    "4.2. PLANO ORÇAMENTÁRIO",
    "GESTAO DE RISCOS",
    "GESTÃO DE RISCOS",
    "PLANO DE GESTAO DE RISCOS",
    "PLANO DE GESTÃO DE RISCOS",
    "4.3. GESTAO DE RISCOS",
    "4.3. GESTÃO DE RISCOS",
    "METAS E PLANOS OBRIGATORIOS",
    "METAS E PLANOS OBRIGATÓRIOS",
    "4. METAS E PLANOS OBRIGATORIOS",
    "4. METAS E PLANOS OBRIGATÓRIOS",
    "CAPACITACAO",
    "CAPACITAÇÃO",
    "5.1. CAPACITACAO",
    "5.1. CAPACITAÇÃO",
    "PLANO DE CAPACITACAO",
    "PLANO DE CAPACITAÇÃO",
    "ACOES ESPECIFICAS",
    "AÇÕES ESPECÍFICAS",
    "5.3. ACOES ESPECIFICAS",
    "5.3. AÇÕES ESPECÍFICAS",
    "5. CAPACITACAO, COMUNICACAO E ACOES ESPECIFICAS",
    "5. CAPACITAÇÃO, COMUNICAÇÃO E AÇÕES ESPECÍFICAS",
    "6. MONITORAMENTO, ANEXOS E REVISAO",
    "6. MONITORAMENTO, ANEXOS E REVISÃO",
    "INDICADORES DE SUCESSO (KPIS)",
}

COLUNAS_METADADOS_IGNORADAS = {
    "UNIDADE",
    "ORGAO",
    "ÓRGÃO",
    "SECRETARIA",
    "SIGLA",
    "ITEM",
    "Nº",
    "N°",
}


def normalizar(texto: str) -> str:
    """Remove espaços extras (inclusive quebras de linha dentro da célula) e coloca em maiúsculas."""
    return " ".join(str(texto).strip().upper().split())


def remover_acentos(texto: str) -> str:
    """Remove acentos mantendo letras e números para comparação flexível."""
    nfkd = unicodedata.normalize("NFKD", normalizar(texto))
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def _resolver_chave_sinonimo(nome_esperado: str) -> str:
    """Retorna a chave canônica em SINONIMOS_SWOT caso nome_esperado seja um sinônimo conhecido."""
    esp_norm = normalizar(nome_esperado)
    esp_sem = remover_acentos(esp_norm)
    if esp_norm in SINONIMOS_SWOT:
        return esp_norm
    for chave_canon, conj_sin in SINONIMOS_SWOT.items():
        if esp_norm in conj_sin or esp_sem in conj_sin or remover_acentos(chave_canon) == esp_sem:
            return chave_canon
    return esp_norm


def limpar_objetivo_estrategico(valor_objetivo: str, valor_id: str = "") -> str:
    """
    Remove o prefixo do ID (ex: 'OE.TIC.1 - ') do texto de 'OBJETIVO ESTRATÉGICO DE TIC',
    pois o código já é colado separadamente no campo 'ID'.
    Exemplo:
      'OE.TIC.1 - Fornecer serviços de TIC eficazes, eficientes e de qualidade'
      -> 'Fornecer serviços de TIC eficazes, eficientes e de qualidade'
    """
    texto = " ".join((valor_objetivo or "").strip().splitlines()).strip()
    if not texto:
        return ""

    id_limpo = " ".join((valor_id or "").strip().splitlines()).strip()
    if id_limpo:
        padrao_id = r"^\s*" + re.escape(id_limpo) + r"(?:\s*[-–—:./]+\s*|\s+)"
        novo_texto = re.sub(padrao_id, "", texto, count=1, flags=re.IGNORECASE).strip()
        if novo_texto and novo_texto != texto:
            return novo_texto

    # Fallback genérico para prefixos do tipo 'OE.TIC.1 - ', 'OE.TIC.01: ', 'OE.1 - ', etc.
    padrao_generico = (
        r"^\s*(?:OE[\s._-]*TIC[\s._-]*\d+|OE[\s._-]*\d+|ID[\s._-]*\d+)"
        r"(?:\s*[-–—:]+\s*|\s+[-–—:]\s*)"
    )
    novo_texto = re.sub(padrao_generico, "", texto, count=1, flags=re.IGNORECASE).strip()
    if novo_texto and novo_texto != texto:
        return novo_texto

    padrao_espaco = r"^\s*OE[\s._-]*TIC[\s._-]*\d+\s+"
    novo_texto = re.sub(padrao_espaco, "", texto, count=1, flags=re.IGNORECASE).strip()
    if novo_texto:
        return novo_texto

    return texto


def _celula_corresponde_coluna(celula: str, nome_esperado: str) -> bool:
    cel_norm = normalizar(celula)
    if not cel_norm or len(cel_norm) > 140:
        return False
    cel_sem = remover_acentos(cel_norm)
    esp_norm = normalizar(nome_esperado)
    esp_sem = remover_acentos(esp_norm)
    chave_canon = _resolver_chave_sinonimo(nome_esperado)

    if chave_canon in SINONIMOS_SWOT:
        sinonimos = SINONIMOS_SWOT[chave_canon]
        if cel_norm in sinonimos or cel_sem in sinonimos:
            return True
        for s in sinonimos:
            if (
                cel_norm.startswith(s + " ")
                or cel_norm.startswith(s + ":")
                or cel_norm.startswith(s + "(")
                or cel_sem.startswith(s + " ")
            ):
                return True
        if chave_canon == "DESCRIÇÃO" and any(
            termo in cel_sem
            for termo in (
                "DESCRICAO SUSCINTA",
                "DESCRICAO SUCINTA",
                "NECESSIDADE PRIORIZADA",
                "CICLO DO PDTIC",
                "DESCRICAO",
            )
        ):
            return True
        if chave_canon == "PRIORIDADE":
            if "CLASSIFICACAO DE PRIORIDADE" in cel_sem:
                return True
            if any(
                termo_num in cel_sem
                for termo_num in (
                    "PONTUACAO",
                    "NOTA",
                    "GRAU",
                    "MEDIA",
                    "RESULTADO",
                    "CALCULO",
                    "INDICE",
                    "FATOR",
                    "PESO",
                    "SOMA",
                )
            ):
                return False
            if any(termo in cel_sem for termo in ("PRIORIDADE", "CLASSIFICACAO")):
                return True
        if chave_canon == "ÁREA DEMANDANTE" and any(
            termo in cel_sem for termo in ("AREA DEMANDANTE", "DEMANDANTE")
        ):
            return True
        if chave_canon == "JUSTIFICATIVA" and "JUSTIFICATIVA" in cel_sem:
            return True
        if chave_canon == "ID":
            if cel_sem in ("ID", "ID OE.TIC", "ID OETIC", "OE.TIC", "ID OE TIC"):
                return True
            if "ID OE.TIC" in cel_sem or "ID OE TIC" in cel_sem or (
                cel_sem.startswith("ID ") and "OE" in cel_sem
            ):
                return True
            return False
        if chave_canon == "OBJETIVO ESTRATÉGICO" and "OBJETIVO ESTRATEGICO" in cel_sem:
            return True
        if chave_canon == "META" and (
            cel_sem in ("META", "METAS", "META DE TIC") or cel_sem.startswith("META ")
        ):
            return True
        if chave_canon == "AÇÃO PRINCIPAL" and any(
            termo in cel_sem
            for termo in (
                "ACAO PRINCIPAL",
                "RESULTADO-CHAVE",
                "RESULTADO CHAVE",
                "KR ESPERADO",
                "(KR)",
            )
        ):
            return True
        if chave_canon == "INDICADOR" and any(
            termo in cel_sem
            for termo in (
                "INDICADOR-CHAVE",
                "INDICADOR CHAVE",
                "INDICADOR DE DESEMPENHO",
                "(KPI)",
                "INDICADOR",
            )
        ):
            return True
        if chave_canon == "PRAZO" and (
            cel_sem in ("PRAZO", "PRAZOS", "QUADRANTE", "QUADRANTE ANO")
            or "PRAZO" in cel_sem
        ):
            return True
        if chave_canon == "RESPONSÁVEL" and (
            "RESPONSAVEL" in cel_sem or "RESPONSAVEIS" in cel_sem
        ):
            return True
        if chave_canon == "FÓRMULA / DESCRIÇÃO" and any(
            termo in cel_sem
            for termo in (
                "METRICA / FORMULA",
                "METRICA/FORMULA",
                "FORMULA / DESCRICAO",
                "FORMULA/DESCRICAO",
                "METRICA",
                "FORMULA",
            )
        ):
            return True
        if chave_canon == "AÇÃO (ID)" and (
            "ACAO (ID)" in cel_sem
            or "ACAO ID" in cel_sem
            or "ID DA ACAO" in cel_sem
            or "ID ACAO" in cel_sem
        ):
            return True
        if chave_canon == "DESCRIÇÃO DO CUSTO" and (
            "CUSTO" in cel_sem or "DESPESA" in cel_sem
        ):
            return True
        if chave_canon == "FONTE DE RECURSO" and "FONTE" in cel_sem:
            return True
        if chave_canon in ("ANO 1", "ANO 2", "ANO 3", "ANO 4"):
            num_ano = chave_canon.split()[-1]
            if (
                f"ANO {num_ano}" in cel_sem
                or f"ANO{num_ano}" in cel_sem
                or f"{num_ano} ANO" in cel_sem
            ):
                return True
        if chave_canon == "RISCO IDENTIFICADO" and (
            "RISCO IDENTIFICADO" in cel_sem
            or cel_sem in ("RISCO", "RISCOS")
            or "DESCRICAO DO RISCO" in cel_sem
        ):
            return True
        if chave_canon == "PROBABILIDADE" and "PROBABILIDADE" in cel_sem:
            return True
        if chave_canon == "IMPACTO" and "IMPACTO" in cel_sem:
            return True
        if chave_canon == "NÍVEL" and (
            "NIVEL" in cel_sem or "GRAU DE RISCO" in cel_sem or "SEVERIDADE" in cel_sem
        ):
            return True
        if chave_canon == "AÇÃO DE MITIGAÇÃO" and (
            "MITIGACAO" in cel_sem or "CONTINGENCIA" in cel_sem or "RESPOSTA AO RISCO" in cel_sem
        ):
            return True
        if chave_canon == "RESPONSÁVEL PELO RISCO" and "RESPONSAVEL" in cel_sem:
            return True
        if chave_canon == "TREINAMENTOS TÉCNICOS" and (
            "TREINAMENTO" in cel_sem or "CAPACITACAO TECNICA" in cel_sem
        ):
            return True
        if chave_canon == "AÇÕES DE RETENÇÃO DE TALENTOS" and (
            "RETENCAO" in cel_sem or "TALENTO" in cel_sem
        ):
            return True
        if chave_canon == "AÇÕES MACRO" and "MACRO" in cel_sem:
            return True
        if chave_canon == "AÇÕES MICRO" and "MICRO" in cel_sem:
            return True
        if chave_canon == "AÇÕES EXTRAS" and ("EXTRA" in cel_sem or "ADICIONA" in cel_sem):
            return True
        return False

    return cel_norm == esp_norm or cel_sem == esp_sem


def _detectar_subcolunas_prazo(
    todas_linhas: List[List[str]],
    indice_linha_cabecalho: int,
    mapa_coluna_para_indice: Dict[str, int],
) -> Tuple[Optional[Tuple[int, int]], bool]:
    """
    Detecta se a coluna 'PRAZO' possui subcolunas duplas ('QUADRANTE' e 'ANO')
    na linha de cabeçalho ou imediatamente abaixo dela (célula mesclada no Excel).
    Retorna ((col_quadrante, col_ano), usou_subcabecalho_abaixo).
    """
    chaves_prazo = [
        k for k in mapa_coluna_para_indice.keys() if _resolver_chave_sinonimo(k) == "PRAZO"
    ]
    if not chaves_prazo:
        return None, False

    chave_prazo = chaves_prazo[0]
    col_prazo = mapa_coluna_para_indice[chave_prazo]
    col_quadrante = -1
    col_ano = -1
    usou_abaixo = False

    if 0 <= indice_linha_cabecalho < len(todas_linhas):
        # Verifica na linha imediatamente abaixo do cabeçalho (onde ficam 'QUADRANTE' e 'ANO')
        if (indice_linha_cabecalho + 1) < len(todas_linhas):
            prox_linha = todas_linhas[indice_linha_cabecalho + 1]
            for idx_c, cel in enumerate(prox_linha):
                cel_sem = remover_acentos(cel)
                if cel_sem in ("QUADRANTE", "QUADRANTES") or "QUADRANTE" in cel_sem:
                    col_quadrante = idx_c
                    usou_abaixo = True
                elif cel_sem in ("ANO", "ANOS") or cel_sem.startswith("ANO "):
                    col_ano = idx_c
                    usou_abaixo = True

        # Verifica na própria linha de cabeçalho caso o usuário tenha copiado a partir da sublinha
        if col_quadrante == -1 and col_ano == -1:
            linha_cab = todas_linhas[indice_linha_cabecalho]
            for idx_c, cel in enumerate(linha_cab):
                cel_sem = remover_acentos(cel)
                if cel_sem in ("QUADRANTE", "QUADRANTES") or "QUADRANTE" in cel_sem:
                    col_quadrante = idx_c
                elif cel_sem in ("ANO", "ANOS") or cel_sem.startswith("ANO "):
                    col_ano = idx_c

    if col_quadrante != -1 and col_ano != -1:
        return (col_quadrante, col_ano), usou_abaixo
    if col_quadrante != -1 and col_ano == -1:
        return (col_quadrante, col_quadrante + 1), usou_abaixo

    # Se 'PRAZO' está em col_prazo e 'RESPONSÁVEL' está em col_prazo + 2, há uma coluna mesclada
    # entre elas (col_prazo = QUADRANTE, col_prazo + 1 = ANO)
    chaves_resp = [
        k for k in mapa_coluna_para_indice.keys() if _resolver_chave_sinonimo(k) == "RESPONSÁVEL"
    ]
    if chaves_resp:
        col_resp = mapa_coluna_para_indice[chaves_resp[0]]
        if col_resp == col_prazo + 2:
            return (col_prazo, col_prazo + 1), usou_abaixo

    return None, usou_abaixo


def _detectar_mapa_por_ambiente_interno_externo(
    todas_linhas: List[List[str]],
) -> tuple[int, Dict[str, int]]:
    """
    Detecta cabeçalhos do tipo:
    ['UNIDADE', 'Ambiente Interno', '', 'Ambiente Externo', '']
    onde 'Ambiente Interno' ocupa (col, col+1) -> (FORÇAS, FRAQUEZAS)
    e 'Ambiente Externo' ocupa (col_ext, col_ext+1) -> (OPORTUNIDADES, AMEAÇAS).
    """
    for idx_linha, linha in enumerate(todas_linhas):
        col_interno = -1
        col_externo = -1
        for idx_col, celula in enumerate(linha):
            cel_sem = remover_acentos(celula)
            if "AMBIENTE INTERNO" in cel_sem and col_interno == -1:
                col_interno = idx_col
            elif "AMBIENTE EXTERNO" in cel_sem and col_externo == -1:
                col_externo = idx_col

        if col_interno != -1 and col_externo != -1 and col_externo >= col_interno + 2:
            return idx_linha, {
                "FORÇAS": col_interno,
                "FRAQUEZAS": col_interno + 1,
                "OPORTUNIDADES": col_externo,
                "AMEAÇAS": col_externo + 1,
            }
    return -1, {}


def extrair_tabela_de_texto_tsv(texto: str, nomes_colunas_esperados: List[str]) -> List[Dict[str, Any]]:
    """
    Converte o texto copiado da planilha (separado por TAB) em uma lista de dicionários.
    Lê todas as linhas da tabela (mesmo quando há linhas em branco, quebras de linha no cabeçalho
    ou células mescladas no Excel) e suporta tanto os cabeçalhos SWOT quanto '3.2. Inventário de necessidades'
    e '4.1. Metas e Ações'.
    """
    texto_limpo = texto.strip("\r\n ")
    if not texto_limpo or "\t" not in texto_limpo:
        return []

    if "AUTOMAÇÃO INTELIGENTE" in texto_limpo or "PS C:\\" in texto_limpo:
        return []

    leitor = csv.reader(io.StringIO(texto_limpo), delimiter="\t")
    todas_linhas = list(leitor)
    linhas_brutas = [linha for linha in todas_linhas if any(celula.strip() for celula in linha)]
    if not linhas_brutas:
        return []

    indice_linha_cabecalho = -1
    usou_linha_abaixo_no_cabecalho = False
    mapa_coluna_para_indice: Dict[str, int] = {}

    # 1. Tenta localizar os cabeçalhos esperados (ex: FORÇAS..AMEAÇAS, DESCRIÇÃO..JUSTIFICATIVA ou ID..RESPONSÁVEL)
    melhor_idx_cabecalho = -1
    melhor_usou_abaixo = False
    melhor_encontrados: Dict[str, int] = {}
    for idx_linha, linha in enumerate(todas_linhas):
        encontrados_nesta_linha: Dict[str, int] = {}
        usou_abaixo_aqui = False
        prox_linha = todas_linhas[idx_linha + 1] if (idx_linha + 1) < len(todas_linhas) else []

        for nome_original in nomes_colunas_esperados:
            for idx_col, celula in enumerate(linha):
                if idx_col in encontrados_nesta_linha.values():
                    continue
                if _celula_corresponde_coluna(celula, nome_original):
                    encontrados_nesta_linha[nome_original] = idx_col
                    break

        # Se algum cabeçalho estiver quebrado/mesclado com a linha imediatamente abaixo
        # (ex: 'DESCRIÇÃO SUSCINTA DE NECESSIDADE' na linha 1 e 'PRIORIZADA NO CICLO DO PDTIC' na linha 2,
        # ou 'CLASSIFICAÇÃO' na linha 1 e 'DE PRIORIDADE' na linha 2)
        if prox_linha:
            for nome_original in nomes_colunas_esperados:
                if nome_original in encontrados_nesta_linha:
                    continue
                max_c = max(len(linha), len(prox_linha))
                for idx_col in range(max_c):
                    if idx_col in encontrados_nesta_linha.values():
                        continue
                    c_topo = linha[idx_col] if idx_col < len(linha) else ""
                    c_baixo = prox_linha[idx_col] if idx_col < len(prox_linha) else ""
                    combinada = f"{c_topo} {c_baixo}".strip()
                    if _celula_corresponde_coluna(c_baixo, nome_original) or _celula_corresponde_coluna(
                        combinada, nome_original
                    ):
                        encontrados_nesta_linha[nome_original] = idx_col
                        usou_abaixo_aqui = True
                        break

        if len(encontrados_nesta_linha) > len(melhor_encontrados):
            melhor_encontrados = encontrados_nesta_linha
            melhor_idx_cabecalho = idx_linha
            melhor_usou_abaixo = usou_abaixo_aqui
            if len(melhor_encontrados) == len(nomes_colunas_esperados):
                break

    if len(melhor_encontrados) >= 1:
        indice_linha_cabecalho = melhor_idx_cabecalho
        usou_linha_abaixo_no_cabecalho = melhor_usou_abaixo
        mapa_coluna_para_indice = dict(melhor_encontrados)

        # Se a linha de cabeçalho encontrou parte das colunas esperadas (ex: 2 ou 3 das 4 colunas
        # do Bloco 2), mapeia as colunas restantes na ordem das células de cabeçalho disponíveis
        if len(mapa_coluna_para_indice) < len(nomes_colunas_esperados):
            linha_cab = todas_linhas[indice_linha_cabecalho]
            prox_cab = (
                todas_linhas[indice_linha_cabecalho + 1]
                if (indice_linha_cabecalho + 1) < len(todas_linhas)
                else []
            )
            max_cols_cab = max(len(linha_cab), len(prox_cab))
            cols_candidatas: List[int] = []
            for idx_c in range(max_cols_cab):
                if idx_c in mapa_coluna_para_indice.values():
                    continue
                txt_c = (
                    (linha_cab[idx_c] if idx_c < len(linha_cab) else "")
                    + " "
                    + (prox_cab[idx_c] if usou_linha_abaixo_no_cabecalho and idx_c < len(prox_cab) else "")
                ).strip()
                txt_sem = remover_acentos(txt_c)
                if not txt_sem or txt_sem in COLUNAS_METADADOS_IGNORADAS or txt_sem in TITULOS_IGNORADOS:
                    continue
                cols_candidatas.append(idx_c)

            faltantes = [n for n in nomes_colunas_esperados if n not in mapa_coluna_para_indice]
            for nome_faltante, idx_cand in zip(faltantes, cols_candidatas):
                mapa_coluna_para_indice[nome_faltante] = idx_cand

    # 2. Se não tinha linha explícita com 'FORÇAS/FRAQUEZAS', verifica cabeçalho 'Ambiente Interno / Externo'
    if indice_linha_cabecalho == -1:
        idx_amb, mapa_amb = _detectar_mapa_por_ambiente_interno_externo(todas_linhas)
        if idx_amb != -1:
            indice_linha_cabecalho = idx_amb
            mapa_coluna_para_indice = mapa_amb

    # 3. Se não é cabeçalho SWOT, verifica se é uma tabela de outra aba/página com cabeçalhos dinâmicos
    if indice_linha_cabecalho == -1:
        texto_up = remover_acentos(texto_limpo)
        parece_swot_sem_cabecalho = any(
            termo in texto_up for termo in ("MATRIZ SWOT", "AMBIENTE INTERNO", "AMBIENTE EXTERNO")
        )

        if not parece_swot_sem_cabecalho:
            for idx_linha, linha in enumerate(todas_linhas):
                nao_vazias = [
                    (idx_c, " ".join(c.strip().splitlines()).strip())
                    for idx_c, c in enumerate(linha)
                    if c.strip() and remover_acentos(c) not in TITULOS_IGNORADOS
                ]
                if len(nao_vazias) >= 2 and all(2 <= len(txt) <= 140 for _, txt in nao_vazias):
                    colunas_dinamicas: Dict[str, int] = {}
                    for idx_c, txt in nao_vazias:
                        if remover_acentos(txt) in COLUNAS_METADADOS_IGNORADAS:
                            continue
                        colunas_dinamicas[txt] = idx_c
                    if len(colunas_dinamicas) >= 2:
                        indice_linha_cabecalho = idx_linha
                        mapa_coluna_para_indice = colunas_dinamicas
                        break

    # Detecta subcolunas duplas de PRAZO ('QUADRANTE' + 'ANO') na aba '4.1. Metas e Ações'
    subcolunas_prazo, sub_prazo_usou_abaixo = _detectar_subcolunas_prazo(
        todas_linhas, indice_linha_cabecalho, mapa_coluna_para_indice
    )
    if sub_prazo_usou_abaixo:
        usou_linha_abaixo_no_cabecalho = True

    # Regra de Memória da Etapa 4 -> Etapa 6:
    # Quando estamos lendo a guia '4.1. Metas e Ações' (que possui 'INDICADOR'), localiza também
    # a coluna 'MÉTRICA / FÓRMULA' na linha de cabeçalho para deixá-la salva em memória para a Etapa 6!
    tem_indicador_no_bloco = any(
        _resolver_chave_sinonimo(k) == "INDICADOR" for k in mapa_coluna_para_indice.keys()
    )
    ja_tem_formula_no_bloco = any(
        _resolver_chave_sinonimo(k) == "FÓRMULA / DESCRIÇÃO" for k in mapa_coluna_para_indice.keys()
    )
    if tem_indicador_no_bloco and not ja_tem_formula_no_bloco and 0 <= indice_linha_cabecalho < len(todas_linhas):
        linha_cab_m = todas_linhas[indice_linha_cabecalho]
        prox_cab_m = (
            todas_linhas[indice_linha_cabecalho + 1]
            if (indice_linha_cabecalho + 1) < len(todas_linhas)
            else []
        )
        max_c_m = max(len(linha_cab_m), len(prox_cab_m))
        for idx_col_m in range(max_c_m):
            if idx_col_m in mapa_coluna_para_indice.values():
                continue
            cel_m1 = linha_cab_m[idx_col_m] if idx_col_m < len(linha_cab_m) else ""
            cel_m2 = prox_cab_m[idx_col_m] if idx_col_m < len(prox_cab_m) else ""
            comb_m = f"{cel_m1} {cel_m2}".strip()
            if (
                _celula_corresponde_coluna(cel_m1, "FÓRMULA / DESCRIÇÃO")
                or _celula_corresponde_coluna(cel_m2, "FÓRMULA / DESCRIÇÃO")
                or _celula_corresponde_coluna(comb_m, "FÓRMULA / DESCRIÇÃO")
            ):
                mapa_coluna_para_indice["FÓRMULA / DESCRIÇÃO"] = idx_col_m
                break

    # Coleta TODAS as linhas de dados abaixo do cabeçalho (sem parar em linhas vazias intermediárias)
    if indice_linha_cabecalho != -1:
        inicio_dados = indice_linha_cabecalho + (2 if usou_linha_abaixo_no_cabecalho else 1)
        linhas_dados = [
            linha
            for linha in todas_linhas[inicio_dados:]
            if any(c.strip() for c in linha)
        ]
    else:
        linhas_dados = []
        for linha in linhas_brutas:
            linha_up = [remover_acentos(c) for c in linha if c.strip()]
            if any(t in TITULOS_IGNORADOS for t in linha_up):
                continue
            linhas_dados.append(linha)

        max_cols = max((len(l) for l in linhas_dados), default=0)
        offset = 1 if max_cols > len(nomes_colunas_esperados) else 0
        for idx, nome_original in enumerate(nomes_colunas_esperados):
            mapa_coluna_para_indice[nome_original] = idx + offset

    # Verificação por conteúdo real das células para 'PRIORIDADE' ('CLASSIFICAÇÃO DE PRIORIDADE'):
    # Se na planilha houver uma coluna numérica de cálculo (ex: '4,33') ao lado da coluna do toggle
    # ('Alta', 'Média', 'Baixa'), garante que 'PRIORIDADE' aponte para a coluna com 'Alta'/'Média'/'Baixa'!
    chaves_prio = [
        k for k in mapa_coluna_para_indice.keys() if _resolver_chave_sinonimo(k) == "PRIORIDADE"
    ]
    if chaves_prio and linhas_dados:
        chave_p = chaves_prio[0]
        cols_fixas = {
            idx_c
            for k, idx_c in mapa_coluna_para_indice.items()
            if k != chave_p
        }
        valores_toggle_prio = {"ALTA", "MEDIA", "BAIXA", "MUITO ALTA", "MUITO BAIXA", "CRITICA"}
        max_c_dados = max((len(l) for l in linhas_dados), default=0)
        melhor_col_toggle = -1
        melhor_contagem_toggle = 0
        for idx_c in range(max_c_dados):
            if idx_c in cols_fixas:
                continue
            contagem = 0
            for l_d in linhas_dados[:30]:
                if idx_c < len(l_d):
                    v_sem = remover_acentos(l_d[idx_c])
                    if v_sem in valores_toggle_prio or (
                        0 < len(v_sem) <= 18
                        and any(p in v_sem.split() for p in ("ALTA", "MEDIA", "BAIXA"))
                    ):
                        contagem += 1
            if contagem > melhor_contagem_toggle:
                melhor_contagem_toggle = contagem
                melhor_col_toggle = idx_c
        if melhor_col_toggle != -1 and melhor_contagem_toggle >= 1:
            mapa_coluna_para_indice[chave_p] = melhor_col_toggle

    nomes_mapeados_norm = {remover_acentos(k) for k in mapa_coluna_para_indice.keys()}
    resultado: List[Dict[str, Any]] = []

    for linha in linhas_dados:
        valores_up = [remover_acentos(c) for c in linha if c.strip()]
        if not valores_up:
            continue
        # Ignora linhas que sejam repetição de títulos ou cabeçalhos (incluindo subcabeçalho QUADRANTE / ANO)
        if all(
            v in TITULOS_IGNORADOS
            or v in COLUNAS_METADADOS_IGNORADAS
            or v in ("QUADRANTE", "ANO", "QUADRANTES", "ANOS")
            for v in valores_up
        ):
            continue
        if sum(1 for v in valores_up if v in nomes_mapeados_norm) >= 2:
            continue

        item: Dict[str, Any] = {}
        for nome_campo, idx_col in mapa_coluna_para_indice.items():
            chave_canon = _resolver_chave_sinonimo(nome_campo)
            if chave_canon == "PRAZO" and subcolunas_prazo is not None:
                col_q, col_a = subcolunas_prazo
                val_q = " ".join((linha[col_q] if col_q < len(linha) else "").strip().splitlines()).strip()
                val_a = " ".join((linha[col_a] if col_a < len(linha) else "").strip().splitlines()).strip()
                if remover_acentos(val_q) in ("QUADRANTE", "QUADRANTES", "PRAZO"):
                    val_q = ""
                if remover_acentos(val_a) in ("ANO", "ANOS", "PRAZO"):
                    val_a = ""
                valor_celula = " ".join(p for p in (val_q, val_a) if p).strip()
            else:
                valor_celula = linha[idx_col].strip() if idx_col < len(linha) else ""
                valor_celula = " ".join(valor_celula.splitlines()).strip()
                if remover_acentos(valor_celula) in SINONIMOS_SWOT.get(chave_canon, set()):
                    valor_celula = ""
            item[nome_campo] = valor_celula

        # Regra especial da aba '4.1. Metas e Ações':
        # Em 'OBJETIVO ESTRATÉGICO', corta o prefixo do ID (ex: 'OE.TIC.1 - ') pois já vai no campo 'ID'
        chave_id_original = next(
            (k for k in item.keys() if _resolver_chave_sinonimo(k) == "ID"),
            None,
        )
        valor_id_linha = item.get(chave_id_original, "") if chave_id_original else ""
        for k_campo in list(item.keys()):
            if _resolver_chave_sinonimo(k_campo) == "OBJETIVO ESTRATÉGICO" and item[k_campo]:
                item[k_campo] = limpar_objetivo_estrategico(item[k_campo], valor_id_linha)

        # Proteção para 'Inventário de necessidades': como 'CLASSIFICAÇÃO DE PRIORIDADE' já vem com
        # um toggle pré-selecionado na planilha, linhas em branco no template que tenham apenas a
        # prioridade preenchida (sem Descrição, Área Demandante nem Justificativa) são ignoradas!
        chaves_item_canon = {
            _resolver_chave_sinonimo(k): v for k, v in item.items()
        }
        if "PRIORIDADE" in chaves_item_canon and (
            "DESCRIÇÃO" in chaves_item_canon or "JUSTIFICATIVA" in chaves_item_canon
        ):
            tem_conteudo_real = any(
                bool(chaves_item_canon.get(ch, "").strip())
                for ch in ("DESCRIÇÃO", "ÁREA DEMANDANTE", "JUSTIFICATIVA")
            )
            if not tem_conteudo_real:
                continue

        # Proteção para '4.1. Metas e Ações': ignora linhas em branco do template que tenham apenas
        # rascunho de ID/Unidade mas nenhum conteúdo em Objetivo Estratégico, Meta, Ação, Indicador ou Responsável
        if "OBJETIVO ESTRATÉGICO" in chaves_item_canon and "META" in chaves_item_canon:
            tem_conteudo_meta = any(
                bool(chaves_item_canon.get(ch, "").strip())
                for ch in (
                    "OBJETIVO ESTRATÉGICO",
                    "META",
                    "AÇÃO PRINCIPAL",
                    "INDICADOR",
                    "PRAZO",
                    "RESPONSÁVEL",
                )
            )
            if not tem_conteudo_meta:
                continue

        if any(item.values()):
            resultado.append(item)

    return resultado


def mesclar_tabelas_por_colunas(
    linhas_base: List[Dict[str, Any]],
    linhas_direita: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    """
    Combina linha a linha as colunas capturadas na esquerda da planilha com as colunas
    capturadas após o scroll para o lado direito na planilha (ex: 'CLASSIFICAÇÃO DE PRIORIDADE',
    'ÁREA DEMANDANTE', 'JUSTIFICATIVA').
    """
    if not linhas_base:
        return list(linhas_direita)
    if not linhas_direita:
        return list(linhas_base)

    total = max(len(linhas_base), len(linhas_direita))
    mescladas: List[Dict[str, Any]] = []
    for i in range(total):
        item_b = dict(linhas_base[i]) if i < len(linhas_base) else {}
        item_d = dict(linhas_direita[i]) if i < len(linhas_direita) else {}
        for k, v in item_d.items():
            if v and not str(item_b.get(k, "")).strip():
                item_b[k] = v
        if any(str(val).strip() for val in item_b.values()):
            mescladas.append(item_b)
    return mescladas


def ler_tabela_copiada_do_clipboard(nomes_colunas_esperados: List[str]) -> List[Dict[str, Any]]:
    """Lê o clipboard atual e extrai a tabela copiada do Excel."""
    texto = ler_do_clipboard()
    return extrair_tabela_de_texto_tsv(texto, nomes_colunas_esperados)


def encontrar_planilha_na_pasta_dados() -> Path | None:
    """Procura se o usuário colocou alguma planilha .xlsx ou .csv na pasta dados/."""
    DADOS_DIR.mkdir(parents=True, exist_ok=True)
    arquivos = [
        p
        for p in DADOS_DIR.iterdir()
        if p.is_file()
        and p.suffix.lower() in [".xlsx", ".xls", ".csv"]
        and p.name not in ["exemplo.xlsx", "progresso_execucao.csv"]
        and not p.name.startswith("~$")
    ]
    return arquivos[0] if arquivos else None


def criar_planilha_exemplo(caminho: Path = ARQUIVO_ENTRADA_PADRAO) -> Path:
    """Cria uma planilha Excel de exemplo caso ainda não exista na pasta dados/."""
    DADOS_DIR.mkdir(parents=True, exist_ok=True)
    if not caminho.exists():
        dados_exemplo = [
            {"id_linha": 1, "codigo": "PRD-001", "descricao": "Teclado Mecânico", "quantidade": 10, "valor": "250,00"},
            {"id_linha": 2, "codigo": "PRD-002", "descricao": "Mouse Óptico", "quantidade": 25, "valor": "89,90"},
            {"id_linha": 3, "codigo": "PRD-003", "descricao": "Monitor 24 Pol", "quantidade": 5, "valor": "899,00"},
        ]
        df = pd.DataFrame(dados_exemplo)
        df.to_excel(caminho, index=False)
    return caminho


def carregar_tabela(caminho: Path = ARQUIVO_ENTRADA_PADRAO, aba: str | int = 0) -> List[Dict[str, Any]]:
    """Lê uma planilha Excel (.xlsx) ou arquivo CSV (.csv)."""
    if not caminho.exists():
        raise FileNotFoundError(f"Arquivo de dados não encontrado: {caminho}")

    sufixo = caminho.suffix.lower()
    if sufixo in [".xlsx", ".xls"]:
        df = pd.read_excel(caminho, sheet_name=aba, dtype=str)
    elif sufixo == ".csv":
        df = pd.read_csv(caminho, dtype=str)
    else:
        raise ValueError(f"Formato de arquivo não suportado: {sufixo}. Use .xlsx ou .csv")

    df = df.fillna("")
    return df.to_dict(orient="records")


def registrar_linha_concluida(indice_linha: int, dados_linha: Dict[str, Any]) -> None:
    """Salva no log de progresso que determinada linha foi concluída com sucesso."""
    DADOS_DIR.mkdir(parents=True, exist_ok=True)
    registro = pd.DataFrame([{"indice_linha": indice_linha, **dados_linha, "status": "OK"}])
    escrever_cabecalho = not ARQUIVO_LOG_PROGRESSO.exists()
    registro.to_csv(
        ARQUIVO_LOG_PROGRESSO,
        mode="a",
        index=False,
        header=escrever_cabecalho,
        encoding="utf-8-sig",
    )

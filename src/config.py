"""
Configurações centralizadas da automação.
Aqui definimos tempos de pausa, caminhos de arquivos e o mapeamento dos campos.
"""

import json
from pathlib import Path
import sys
from typing import Any, Dict, List, Optional, Tuple
import unicodedata

try:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE_DIR = Path(__file__).resolve().parent.parent
DADOS_DIR = BASE_DIR / "dados"
DIRETORIO_DADOS = DADOS_DIR
ARQUIVO_ENTRADA_PADRAO = DADOS_DIR / "exemplo.xlsx"
ARQUIVO_LOG_PROGRESSO = DADOS_DIR / "progresso_execucao.csv"
ARQUIVO_MAPEAMENTO = DADOS_DIR / "mapeamento_campos.json"

TEMPO_CONTAGEM_INICIAL = 5
ESPERA_CURTA = 0.35
ESPERA_MEDIA = 0.75
ESPERA_APOS_ADICIONAR = 1.45  # Pausa após clicar em "Adicionar" para renderização DOM da nova linha

TECLA_PARADA_EMERGENCIA = "esc"
MODO_SIMULACAO = False
DESLOCAMENTO_Y_POR_LINHA = 0

# Modelos de páginas do sistema (padrões sugeridos por página, permitindo criar/editar livremente)
PAGINAS_PADRAO_SISTEMA: Dict[str, Dict[str, Any]] = {
    "1. Identificação e Aprovação": {
        "etapa": 1,
        "caixas_sugeridas": [
            {"nome": "VIGÊNCIA", "tipo": "caixa_texto", "tem_adicionar": False},
            {"nome": "VERSÃO", "tipo": "caixa_texto", "tem_adicionar": False},
            {"nome": "ANO", "tipo": "caixa_texto", "tem_adicionar": False},
            {"nome": "CIDADE", "tipo": "caixa_texto", "tem_adicionar": False},
            {"nome": "TEXTO DE APROVAÇÃO", "tipo": "caixa_texto", "tem_adicionar": False},
            {"nome": "AUTORIDADE MÁXIMA - NOME", "tipo": "caixa_texto", "tem_adicionar": False},
            {"nome": "AUTORIDADE MÁXIMA - CARGO", "tipo": "caixa_texto", "tem_adicionar": False},
            {"nome": "AUTORIDADE MÁXIMA - DATA", "tipo": "caixa_texto", "tem_adicionar": False},
            {"nome": "GESTOR DE TI - NOME", "tipo": "caixa_texto", "tem_adicionar": False},
            {"nome": "GESTOR DE TI - CARGO", "tipo": "caixa_texto", "tem_adicionar": False},
            {"nome": "GESTOR DE TI - DATA", "tipo": "caixa_texto", "tem_adicionar": False},
            {"nome": "COMITÊ - NOME", "tipo": "caixa_texto", "tem_adicionar": True},
            {"nome": "COMITÊ - CARGO/REPRESENTAÇÃO", "tipo": "caixa_texto", "tem_adicionar": False},
        ],
    },
    "2. Fundamentação": {
        "etapa": 2,
        "caixas_sugeridas": [
            {"nome": "SUMÁRIO EXECUTIVO", "tipo": "caixa_texto", "tem_adicionar": False},
            {"nome": "SIGLA", "tipo": "caixa_texto", "tem_adicionar": True},
            {"nome": "SIGNIFICADO", "tipo": "caixa_texto", "tem_adicionar": False},
            {"nome": "JUSTIFICATIVA", "tipo": "caixa_texto", "tem_adicionar": False},
            {"nome": "OBJETIVO", "tipo": "caixa_texto", "tem_adicionar": False},
            {"nome": "ESCOPO", "tipo": "caixa_texto", "tem_adicionar": False},
            {"nome": "ABRANGÊNCIA", "tipo": "caixa_texto", "tem_adicionar": False},
            {"nome": "PETIC", "tipo": "caixa_texto", "tem_adicionar": False},
            {"nome": "PPA", "tipo": "caixa_texto", "tem_adicionar": False},
            {"nome": "CT-IPE", "tipo": "caixa_texto", "tem_adicionar": False},
            {"nome": "TÍTULO DA REFERÊNCIA", "tipo": "caixa_texto", "tem_adicionar": True},
        ],
    },
    "3. Diagnóstico": {
        "etapa": 3,
        "caixas_sugeridas": [
            {"nome": "FORÇAS", "tipo": "caixa_texto", "tem_adicionar": True},
            {"nome": "FRAQUEZAS", "tipo": "caixa_texto", "tem_adicionar": True},
            {"nome": "OPORTUNIDADES", "tipo": "caixa_texto", "tem_adicionar": True},
            {"nome": "AMEAÇAS", "tipo": "caixa_texto", "tem_adicionar": True},
            {"nome": "DESCRIÇÃO", "tipo": "caixa_texto", "tem_adicionar": True, "grupo": "INVENTÁRIO DE NECESSIDADES"},
            {
                "nome": "PRIORIDADE",
                "tipo": "barra_suspensa",
                "tem_adicionar": False,
                "opcoes_barra_suspensa": ["Alta", "Média", "Baixa"],
                "grupo": "INVENTÁRIO DE NECESSIDADES",
            },
            {"nome": "ÁREA DEMANDANTE", "tipo": "caixa_texto", "tem_adicionar": False, "grupo": "INVENTÁRIO DE NECESSIDADES"},
            {"nome": "JUSTIFICATIVA", "tipo": "caixa_texto", "tem_adicionar": False, "grupo": "INVENTÁRIO DE NECESSIDADES"},
        ],
    },
    "4. Metas e Planos Obrigatórios": {
        "etapa": 4,
        "caixas_sugeridas": [
            {"nome": "ID", "tipo": "caixa_texto", "tem_adicionar": True, "grupo": "METAS E AÇÕES"},
            {"nome": "OBJETIVO ESTRATÉGICO", "tipo": "caixa_texto", "tem_adicionar": False, "grupo": "METAS E AÇÕES"},
            {"nome": "META", "tipo": "caixa_texto", "tem_adicionar": False, "grupo": "METAS E AÇÕES"},
            {"nome": "AÇÃO PRINCIPAL", "tipo": "caixa_texto", "tem_adicionar": False, "grupo": "METAS E AÇÕES"},
            {"nome": "INDICADOR", "tipo": "caixa_texto", "tem_adicionar": False, "grupo": "METAS E AÇÕES"},
            {"nome": "PRAZO", "tipo": "caixa_texto", "tem_adicionar": False, "grupo": "METAS E AÇÕES"},
            {"nome": "RESPONSÁVEL", "tipo": "caixa_texto", "tem_adicionar": False, "grupo": "METAS E AÇÕES"},
            {"nome": "AÇÃO (ID)", "tipo": "caixa_texto", "tem_adicionar": True, "grupo": "PLANO ORÇAMENTÁRIO"},
            {"nome": "DESCRIÇÃO DO CUSTO", "tipo": "caixa_texto", "tem_adicionar": False, "grupo": "PLANO ORÇAMENTÁRIO"},
            {"nome": "FONTE DE RECURSO", "tipo": "caixa_texto", "tem_adicionar": False, "grupo": "PLANO ORÇAMENTÁRIO"},
            {"nome": "ANO 1", "tipo": "caixa_texto", "tem_adicionar": False, "grupo": "PLANO ORÇAMENTÁRIO"},
            {"nome": "ANO 2", "tipo": "caixa_texto", "tem_adicionar": False, "grupo": "PLANO ORÇAMENTÁRIO"},
            {"nome": "ANO 3", "tipo": "caixa_texto", "tem_adicionar": False, "grupo": "PLANO ORÇAMENTÁRIO"},
            {"nome": "ANO 4", "tipo": "caixa_texto", "tem_adicionar": False, "grupo": "PLANO ORÇAMENTÁRIO"},
            {"nome": "TOTAL", "tipo": "caixa_texto", "tem_adicionar": False, "grupo": "PLANO ORÇAMENTÁRIO"},
            {"nome": "RISCO IDENTIFICADO", "tipo": "caixa_texto", "tem_adicionar": True, "grupo": "GESTÃO DE RISCOS"},
            {"nome": "PROBABILIDADE", "tipo": "caixa_texto", "tem_adicionar": False, "grupo": "GESTÃO DE RISCOS"},
            {"nome": "IMPACTO", "tipo": "caixa_texto", "tem_adicionar": False, "grupo": "GESTÃO DE RISCOS"},
            {"nome": "NÍVEL", "tipo": "caixa_texto", "tem_adicionar": False, "grupo": "GESTÃO DE RISCOS"},
            {"nome": "AÇÃO DE MITIGAÇÃO", "tipo": "caixa_texto", "tem_adicionar": False, "grupo": "GESTÃO DE RISCOS"},
            {"nome": "RESPONSÁVEL PELO RISCO", "tipo": "caixa_texto", "tem_adicionar": False, "grupo": "GESTÃO DE RISCOS"},
        ],
    },
    "5. Capacitação, Comunicação e Ações Específicas": {
        "etapa": 5,
        "caixas_sugeridas": [
            {"nome": "TREINAMENTOS TÉCNICOS", "tipo": "caixa_texto", "tem_adicionar": True},
            {"nome": "AÇÕES DE RETENÇÃO DE TALENTOS", "tipo": "caixa_texto", "tem_adicionar": True},
            {"nome": "PÚBLICOS-ALVO", "tipo": "caixa_texto", "tem_adicionar": True},
            {"nome": "CANAIS", "tipo": "caixa_texto", "tem_adicionar": True},
            {"nome": "PERIODICIDADE", "tipo": "caixa_texto", "tem_adicionar": False},
            {"nome": "RESPONSÁVEL", "tipo": "caixa_texto", "tem_adicionar": False},
            {"nome": "AÇÕES MACRO", "tipo": "caixa_texto", "tem_adicionar": True},
            {"nome": "AÇÕES MICRO", "tipo": "caixa_texto", "tem_adicionar": True},
            {"nome": "AÇÕES EXTRAS", "tipo": "caixa_texto", "tem_adicionar": True},
        ],
    },
    "6. Monitoramento, Anexos e Revisão": {
        "etapa": 6,
        "caixas_sugeridas": [
            {"nome": "INDICADOR", "tipo": "caixa_texto", "tem_adicionar": True, "grupo": "INDICADORES DE SUCESSO (KPIS)"},
            {"nome": "FÓRMULA / DESCRIÇÃO", "tipo": "caixa_texto", "tem_adicionar": False, "grupo": "INDICADORES DE SUCESSO (KPIS)"},
            {"nome": "RITOS DE MONITORAMENTO", "tipo": "caixa_texto", "tem_adicionar": True},
            {"nome": "ANEXOS", "tipo": "caixa_texto", "tem_adicionar": True},
        ],
    },
}

# Sequência oficial das fases/etapas do sistema que serão preenchidas pelo robô
SEQUENCIA_FASES_SISTEMA: List[str] = [
    "3. Diagnóstico",
    "4. Metas e Planos Obrigatórios",
    "5. Capacitação, Comunicação e Ações Específicas",
    "6. Monitoramento, Anexos e Revisão",
]

# Mapeamento padrão (usado caso você ainda não tenha rodado o `treinador.py`)
MAPEAMENTO_PADRAO: Dict[str, Any] = {
    "modo_origem": "planilha_aberta_ctrl_c",
    "planilha": [-1649, 373],
    "coord_toggle_todas_guias": None,
    "coord_botao_proxima_etapa": None,
    "pagina_ativa": "3. Diagnóstico",
    "campos_detalhados": {},
    "padroes_paginas": {},
    "deslocamento_y_por_linha": 0,
}


def _norm_chave(texto: str) -> str:
    sem_acento = "".join(
        c for c in unicodedata.normalize("NFD", str(texto)) if unicodedata.category(c) != "Mn"
    )
    return sem_acento.strip().upper()


def gerar_passos_padrao_da_pagina(info_pagina: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    Gera ou atualiza a lista ordenada `passos_pagina` de uma página a partir de seus `campos`
    e passos já configurados:
    - Mantém grupos codependentes (`tipo_passo == "grupo_codependente"`) já ensinados pelo usuário.
    - Para campos que ainda não estão em nenhum passo, cria passos individuais (`coluna_independente`
      ou `campo_unico`) na ordem em que foram mapeados.
    """
    campos_pag = info_pagina.get("campos") or {}
    passos_existentes = info_pagina.get("passos_pagina")
    passos_validos: List[Dict[str, Any]] = []
    campos_ja_incluidos = set()

    if isinstance(passos_existentes, list):
        for p in passos_existentes:
            if not isinstance(p, dict):
                continue
            tipo_p = str(p.get("tipo_passo") or "coluna_independente")
            lista_c = [str(c) for c in (p.get("campos") or []) if str(c) in campos_pag]
            if not lista_c and p.get("campo") and str(p.get("campo")) in campos_pag:
                lista_c = [str(p.get("campo"))]
            if not lista_c:
                continue
            passo_limpo = dict(p)
            passo_limpo["tipo_passo"] = tipo_p
            passo_limpo["campos"] = lista_c
            passo_limpo.setdefault("nome_passo", str(p.get("nome_passo") or " + ".join(lista_c)))
            passos_validos.append(passo_limpo)
            for c_nome in lista_c:
                campos_ja_incluidos.add(c_nome)

    # Agrupa campos que têm `grupo_codependente` definido nos seus metadados mas ainda não estavam em `passos_pagina`
    grupos_pendentes: Dict[str, List[str]] = {}
    for nome_c, meta_c in campos_pag.items():
        if nome_c in campos_ja_incluidos or not isinstance(meta_c, dict):
            continue
        if meta_c.get("tipo") == "bloqueada" or meta_c.get("somente_bloqueio"):
            continue
        nome_grupo = str(meta_c.get("grupo_codependente") or "").strip()
        if nome_grupo:
            grupos_pendentes.setdefault(nome_grupo, []).append(nome_c)
            campos_ja_incluidos.add(nome_c)

    for nome_grupo, lista_c in grupos_pendentes.items():
        btn_grupo = None
        passo_y_g = None
        for c_nome in lista_c:
            meta_c = campos_pag.get(c_nome) or {}
            if meta_c.get("botao_adicionar") and not btn_grupo:
                btn_grupo = meta_c.get("botao_adicionar")
            if meta_c.get("passo_y_grupo") and not passo_y_g:
                passo_y_g = meta_c.get("passo_y_grupo")
        passos_validos.append(
            {
                "tipo_passo": "grupo_codependente",
                "nome_passo": nome_grupo,
                "campos": lista_c,
                "botao_adicionar": btn_grupo,
                "passo_y_grupo": passo_y_g,
            }
        )

    # Adiciona os demais campos independentes na ordem em que estão no dicionário `campos`
    for nome_c, meta_c in campos_pag.items():
        if nome_c in campos_ja_incluidos or not isinstance(meta_c, dict):
            continue
        if meta_c.get("tipo") == "bloqueada" or meta_c.get("somente_bloqueio"):
            continue
        tem_add = bool(meta_c.get("tem_adicionar", bool(meta_c.get("botao_adicionar"))))
        passos_validos.append(
            {
                "tipo_passo": "coluna_independente" if tem_add else "campo_unico",
                "nome_passo": nome_c,
                "campos": [nome_c],
            }
        )
        campos_ja_incluidos.add(nome_c)

    return passos_validos


def sincronizar_padroes_paginas(dados: Dict[str, Any]) -> Dict[str, Any]:
    """
    Garante que o JSON de mapeamento possua a estrutura `padroes_paginas` sincronizada
    com `campos_detalhados` / `campos` e `passos_pagina`, preservando todos os campos e regras
    já calibrados pelo usuário.
    """
    if not isinstance(dados, dict):
        dados = dict(MAPEAMENTO_PADRAO)

    padroes = dados.get("padroes_paginas")
    if not isinstance(padroes, dict):
        padroes = {}
        dados["padroes_paginas"] = padroes

    pagina_ativa = str(dados.get("pagina_ativa") or "3. Diagnóstico")
    dados["pagina_ativa"] = pagina_ativa

    # Migra campos já existentes na raiz para o padrão da página ativa caso ainda não estejam lá
    campos_raiz = dados.get("campos_detalhados") or {}
    if campos_raiz and pagina_ativa not in padroes:
        campos_migrados: Dict[str, Any] = {}
        for nome_c, info_c in campos_raiz.items():
            if isinstance(info_c, dict):
                item_copia = dict(info_c)
                item_copia.setdefault("tipo", "caixa_texto")
                item_copia.setdefault("tem_adicionar", bool(info_c.get("botao_adicionar")))
                campos_migrados[nome_c] = item_copia
        padroes[pagina_ativa] = {
            "nome_pagina": pagina_ativa,
            "campos": campos_migrados,
            "caixas_bloqueadas": [],
            "passos_pagina": [],
        }

    # Sincroniza `passos_pagina` em todas as páginas cadastradas
    for nome_p, info_p in padroes.items():
        if isinstance(info_p, dict):
            info_p.setdefault("nome_pagina", nome_p)
            info_p.setdefault("campos", {})
            info_p.setdefault("caixas_bloqueadas", [])
            info_p["passos_pagina"] = gerar_passos_padrao_da_pagina(info_p)

    # Mantém `campos_detalhados` e `campos` na raiz atualizados com a página ativa para retrocompatibilidade
    if pagina_ativa in padroes and isinstance(padroes[pagina_ativa], dict):
        campos_pag = padroes[pagina_ativa].get("campos") or {}
        if campos_pag:
            dados["campos_detalhados"] = campos_pag
            dados["campos"] = {
                k: v.get("sistema_campo_linha1", [910, 488])
                for k, v in campos_pag.items()
                if isinstance(v, dict) and v.get("sistema_campo_linha1")
            }

    return dados


def identificar_padrao_pagina_pelas_colunas(
    nomes_colunas_planilha: List[str],
    dados: Dict[str, Any],
    pagina_forcada: Optional[str] = None,
) -> Tuple[str, Dict[str, Any]]:
    """
    Identifica automaticamente qual Padrão de Página (`padroes_paginas`) corresponde às colunas
    da planilha atual (ou usa `pagina_forcada` se informado).
    Retorna `(nome_pagina, dict_padrao_pagina)`.
    """
    dados = sincronizar_padroes_paginas(dados)
    padroes: Dict[str, Any] = dados.get("padroes_paginas") or {}

    if pagina_forcada:
        chave_f = _norm_chave(pagina_forcada)
        for nome_p, info_p in padroes.items():
            if _norm_chave(nome_p) == chave_f or chave_f in _norm_chave(nome_p):
                return nome_p, info_p

    colunas_norm = [_norm_chave(c) for c in nomes_colunas_planilha if str(c).strip()]
    melhor_pagina = str(dados.get("pagina_ativa") or "3. Diagnóstico")
    melhor_score = -1

    for nome_p, info_p in padroes.items():
        if not isinstance(info_p, dict):
            continue
        campos_p = info_p.get("campos") or {}
        chaves_campos_p = [_norm_chave(k) for k in campos_p.keys()]
        score = 0
        for col_n in colunas_norm:
            for ch_p in chaves_campos_p:
                if col_n == ch_p or (len(ch_p) >= 4 and (ch_p in col_n or col_n in ch_p)):
                    score += 10
                    break
        if score > melhor_score and score > 0:
            melhor_score = score
            melhor_pagina = nome_p

    if melhor_pagina in padroes and isinstance(padroes[melhor_pagina], dict):
        return melhor_pagina, padroes[melhor_pagina]

    # Fallback para os campos da raiz
    info_fallback = {
        "nome_pagina": melhor_pagina,
        "campos": dados.get("campos_detalhados") or {},
        "caixas_bloqueadas": [],
        "passos_pagina": [],
    }
    info_fallback["passos_pagina"] = gerar_passos_padrao_da_pagina(info_fallback)
    return melhor_pagina, info_fallback


def carregar_mapeamento() -> Dict[str, Any]:
    """
    Carrega o mapeamento de campos salvo pelo `treinador.py` (dados/mapeamento_campos.json).
    Se o arquivo ainda não existir, cria e retorna o mapeamento padrão.
    """
    DADOS_DIR.mkdir(parents=True, exist_ok=True)
    if not ARQUIVO_MAPEAMENTO.exists():
        salvar_mapeamento(MAPEAMENTO_PADRAO)
        return dict(MAPEAMENTO_PADRAO)

    with open(ARQUIVO_MAPEAMENTO, "r", encoding="utf-8") as f:
        dados = json.load(f)
    return sincronizar_padroes_paginas(dados)


def salvar_mapeamento(dados: Dict[str, Any]) -> None:
    """Salva o dicionário de coordenadas dos campos, páginas e botões Adicionar em JSON."""
    DADOS_DIR.mkdir(parents=True, exist_ok=True)
    dados = sincronizar_padroes_paginas(dados)
    with open(ARQUIVO_MAPEAMENTO, "w", encoding="utf-8") as f:
        json.dump(dados, f, indent=4, ensure_ascii=False)


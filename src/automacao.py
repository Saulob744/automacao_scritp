"""
src/automacao.py
Fluxo de automação dinâmica:
1. Seleciona tudo e copia na planilha aberta (Ctrl+A -> Ctrl+C).
2. Usa o conjunto de coordenadas de parâmetro + Ctrl+F + HTML + Visão OpenCV para localizar cada campo.
3. Cola linha por linha da tabela, clicando em 'Adicionar' para abrir e preencher as caixas seguintes.
"""

import json
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from src.config import (
    ARQUIVO_LOG_PROGRESSO,
    DIRETORIO_DADOS,
    ESPERA_APOS_ADICIONAR,
    ESPERA_CURTA,
    SEQUENCIA_FASES_SISTEMA,
    carregar_mapeamento,
    identificar_padrao_pagina_pelas_colunas,
    salvar_mapeamento,
)
from src.driver_windows import (
    aguardar_tecla_f8,
    atalho_teclado,
    copiar_para_clipboard,
    iniciar_monitor_parada_emergencia,
    ler_do_clipboard,
    mover_e_clicar,
    pressionar_tecla_unica,
    rolar_pagina,
    tecla_pressionada,
)
from src.leitor_dados import (
    _celula_corresponde_coluna,
    _resolver_chave_sinonimo,
    extrair_tabela_de_texto_tsv,
    mesclar_tabelas_por_colunas,
    normalizar,
    registrar_linha_concluida,
    remover_acentos,
)
from src.localizador_visual import (
    _medir_scroll_por_template,
    ajustar_coordenada_para_caixa_vazia,
    capturar_tela_sistema,
    centralizar_y_na_caixa_branca,
    clicar_adicionar_e_avancar_caixa,
    clicar_adicionar_e_avancar_grupo_codependente,
    clicar_botao_proxima_etapa_no_sistema,
    detectar_botoes_remover_da_coluna,
    eh_caixa_de_texto_vazia,
    ir_para_topo_da_pagina,
    localizar_card_com_ctrl_f_e_visao,
    obter_mapa_html_campos,
    obter_parametros_calibrados,
    rolar_pagina_ate_o_final,
    rolar_planilha_para_direita,
    selecionar_guia_planilha_via_toggle_e_nome,
    selecionar_guia_wizard_no_sistema,
    selecionar_opcao_barra_suspensa,
)

COLUNAS_PADRAO_SWOT = ["FORÇAS", "FRAQUEZAS", "OPORTUNIDADES", "AMEAÇAS"]
COLUNAS_INVENTARIO_NECESSIDADES = ["DESCRIÇÃO", "PRIORIDADE", "ÁREA DEMANDANTE", "JUSTIFICATIVA"]
COLUNAS_METAS_ACOES = [
    "ID",
    "OBJETIVO ESTRATÉGICO",
    "META",
    "AÇÃO PRINCIPAL",
    "INDICADOR",
    "PRAZO",
    "RESPONSÁVEL",
]
COLUNAS_PLANO_ORCAMENTARIO = [
    "AÇÃO (ID)",
    "DESCRIÇÃO DO CUSTO",
    "FONTE DE RECURSO",
    "ANO 1",
    "ANO 2",
    "ANO 3",
    "ANO 4",
]
COLUNAS_GESTAO_RISCOS = [
    "RISCO IDENTIFICADO",
    "PROBABILIDADE",
    "IMPACTO",
    "NÍVEL",
    "AÇÃO DE MITIGAÇÃO",
    "RESPONSÁVEL PELO RISCO",
]
COLUNAS_CAPACITACAO_5_1 = [
    "TREINAMENTOS TÉCNICOS",
    "AÇÕES DE RETENÇÃO DE TALENTOS",
]
COLUNAS_ACOES_ESPECIFICAS_5_3 = [
    "AÇÕES MACRO",
    "AÇÕES MICRO",
    "AÇÕES EXTRAS",
]
COLUNAS_MONITORAMENTO_KPIS_6 = [
    "INDICADOR",
    "FÓRMULA / DESCRIÇÃO",
]

# Cache em memória dos dados coletados na Etapa 4 ('4.1. Metas e Ações'),
# reaproveitado diretamente na Etapa 6 ('Indicadores de sucesso (KPIs)') sem voltar na planilha!
CACHE_MEMORIA_ETAPA_4: List[Dict[str, Any]] = []
ARQUIVO_CACHE_ETAPA_4: Path = DIRETORIO_DADOS / "cache_etapa_4.json"

# Limite configurável de linhas/itens a serem copiados e colados por bloco/coluna durante os testes
LIMITE_LINHAS_TESTE: Optional[int] = None


def definir_limite_linhas_teste(limite: Optional[int]) -> None:
    """Define o limite máximo de linhas/itens a serem copiados e colados de cada bloco/coluna no teste."""
    global LIMITE_LINHAS_TESTE
    if limite is not None and int(limite) > 0:
        LIMITE_LINHAS_TESTE = int(limite)
        print(
            f"🔢 [MODO DE TESTE COM LIMITE] Serão copiadas/coladas no máximo "
            f"{LIMITE_LINHAS_TESTE} linha(s) de cada bloco/coluna!"
        )
    else:
        LIMITE_LINHAS_TESTE = None


def obter_limite_linhas_teste() -> Optional[int]:
    """Retorna o limite de linhas configurado para o teste atual (ou None para copiar todas)."""
    return LIMITE_LINHAS_TESTE


def salvar_dados_etapa_4_em_memoria(
    linhas_etapa_4: List[Dict[str, Any]],
    mapeamento: Optional[Dict[str, Any]] = None,
) -> List[Dict[str, Any]]:
    """
    Armazena em memória (CACHE_MEMORIA_ETAPA_4) e em arquivo local temporário (dados/cache_etapa_4.json)
    a matriz coletada na Etapa 4 ('4.1. Metas e Ações') para reaproveitamento direto na Etapa 6:
    - 'INDICADOR-CHAVE DE DESEMPENHO (KPI)' <-> 'INDICADOR'
    - 'MÉTRICA / FÓRMULA' <-> 'FÓRMULA / DESCRIÇÃO'
    """
    global CACHE_MEMORIA_ETAPA_4
    matriz_normalizada: List[Dict[str, Any]] = []
    for linha in linhas_etapa_4 or []:
        copia = dict(linha)
        val_ind = (
            copia.get("INDICADOR")
            or copia.get("INDICADOR-CHAVE DE DESEMPENHO (KPI)")
            or ""
        )
        val_form = (
            copia.get("FÓRMULA / DESCRIÇÃO")
            or copia.get("_FORMULA_DESCRICAO_CACHE")
            or copia.get("MÉTRICA / FÓRMULA")
            or copia.get("METRICA / FORMULA")
            or ""
        )
        if val_ind:
            copia["INDICADOR"] = str(val_ind).strip()
            copia["INDICADOR-CHAVE DE DESEMPENHO (KPI)"] = str(val_ind).strip()
        if val_form:
            copia["FÓRMULA / DESCRIÇÃO"] = str(val_form).strip()
            copia["MÉTRICA / FÓRMULA"] = str(val_form).strip()
        matriz_normalizada.append(copia)

    CACHE_MEMORIA_ETAPA_4 = matriz_normalizada
    if matriz_normalizada:
        try:
            ARQUIVO_CACHE_ETAPA_4.parent.mkdir(parents=True, exist_ok=True)
            ARQUIVO_CACHE_ETAPA_4.write_text(
                json.dumps(matriz_normalizada, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception:
            pass
    print(
        f"🧠 [MEMÓRIA ETAPA 4 -> ETAPA 6] {len(CACHE_MEMORIA_ETAPA_4)} registro(s) de '4.1. Metas e Ações' "
        "(incluindo 'INDICADOR-CHAVE DE DESEMPENHO (KPI)' e 'MÉTRICA / FÓRMULA') salvos na memória!"
    )
    return CACHE_MEMORIA_ETAPA_4


def obter_dados_etapa_4_da_memoria(
    mapeamento: Optional[Dict[str, Any]] = None,
) -> List[Dict[str, Any]]:
    """
    Recupera os dados da Etapa 4 armazenados em memória (CACHE_MEMORIA_ETAPA_4) ou no cache local
    para preencher a Etapa 6 ('Indicadores de sucesso (KPIs)') sem precisar voltar à planilha.
    """
    if CACHE_MEMORIA_ETAPA_4:
        return list(CACHE_MEMORIA_ETAPA_4)
    if ARQUIVO_CACHE_ETAPA_4.exists():
        try:
            dados_arquivo = json.loads(ARQUIVO_CACHE_ETAPA_4.read_text(encoding="utf-8"))
            if isinstance(dados_arquivo, list) and dados_arquivo:
                return list(dados_arquivo)
        except Exception:
            pass
    if isinstance(mapeamento, dict):
        cache_salvo = mapeamento.get("cache_memoria_etapa_4")
        if isinstance(cache_salvo, list) and cache_salvo:
            return list(cache_salvo)
    return []


def verificar_aborto() -> bool:
    if tecla_pressionada("esc"):
        print("\n🛑 [ABORTADO] Tecla ESC detectada! Interrompendo automação imediatamente.")
        return True
    return False


def _colar_texto_no_campo(valor: str) -> None:
    """Copia o valor para o clipboard, limpa o campo focado e cola com Ctrl+V."""
    copiar_para_clipboard(str(valor))
    time.sleep(0.08)
    atalho_teclado("ctrl", "a")
    time.sleep(0.05)
    pressionar_tecla_unica("backspace")
    time.sleep(0.05)
    atalho_teclado("ctrl", "v")


def buscar_valor_da_coluna(linha: Dict[str, Any], nome_campo: str) -> Optional[str]:
    """
    Busca o valor correspondente ao campo na linha da planilha, reconhecendo sinônimos
    como:
    - 'DESCRIÇÃO' <-> 'DESCRIÇÃO SUSCINTA DE NECESSIDADE PRIORIZADA NO CICLO DO PDTIC'
    - 'PRIORIDADE' <-> 'CLASSIFICAÇÃO DE PRIORIDADE'
    - 'ÁREA DEMANDANTE' <-> 'ÁREA DEMANDANTE'
    - 'JUSTIFICATIVA' <-> 'JUSTIFICATIVA'
    """
    if nome_campo in linha and str(linha[nome_campo]).strip():
        return str(linha[nome_campo]).strip()

    alvo_norm = remover_acentos(nome_campo)
    alvo_singular = alvo_norm[:-1] if alvo_norm.endswith("S") else alvo_norm
    chave_canon_alvo = _resolver_chave_sinonimo(nome_campo)

    for col_planilha, valor in linha.items():
        if str(col_planilha).startswith("_"):
            continue
        if not str(valor or "").strip():
            continue
        col_norm = remover_acentos(col_planilha)
        col_singular = col_norm[:-1] if col_norm.endswith("S") else col_norm
        if (
            col_norm == alvo_norm
            or col_singular == alvo_singular
            or _resolver_chave_sinonimo(col_planilha) == chave_canon_alvo
            or _celula_corresponde_coluna(col_planilha, nome_campo)
        ):
            return str(valor).strip()

    return None


def _buscar_info_campo_calibrado(
    nome_campo: str,
    campos_detalhados: Dict[str, Any],
) -> Dict[str, Any]:
    """Retorna o dicionário de configuração de `nome_campo` dentro de `campos_detalhados`."""
    if nome_campo in campos_detalhados and isinstance(campos_detalhados[nome_campo], dict):
        return campos_detalhados[nome_campo]
    alvo = remover_acentos(nome_campo)
    alvo_s = alvo[:-1] if alvo.endswith("S") else alvo
    chave_canon = _resolver_chave_sinonimo(nome_campo)
    for k_c, v_c in (campos_detalhados or {}).items():
        if not isinstance(v_c, dict):
            continue
        kn = remover_acentos(k_c)
        ks = kn[:-1] if kn.endswith("S") else kn
        if kn == alvo or ks == alvo_s or _resolver_chave_sinonimo(k_c) == chave_canon:
            return v_c
    return {}


def _extrair_tabela_auto_detectando_bloco(
    texto_copiado: str,
    mapeamento: Dict[str, Any],
) -> List[Dict[str, Any]]:
    """
    Identifica automaticamente a qual aba/bloco da planilha pertence o texto TSV copiado
    ('4.1. Metas e Ações', '4.2. Plano Orçamentário', '4.3. Gestão de Riscos',
    '5.1. Capacitação', '5.3. Ações Específicas', '3.2. Inventário de necessidades'
    ou '3.1. Análise SWOT') e extrai as colunas correspondentes sem misturar colunas de páginas diferentes.
    """
    texto_up = remover_acentos(texto_copiado or "")
    if any(
        termo in texto_up
        for termo in (
            "METAS E ACOES",
            "OBJETIVO ESTRATEGICO",
            "RESULTADO-CHAVE",
            "RESULTADO CHAVE",
            "INDICADOR-CHAVE",
            "INDICADOR CHAVE",
            "OE.TIC",
        )
    ):
        linhas = extrair_tabela_de_texto_tsv(texto_copiado, list(COLUNAS_METAS_ACOES))
        if linhas:
            return linhas

    if any(
        termo in texto_up
        for termo in (
            "PLANO ORCAMENTARIO",
            "DESCRICAO DO CUSTO",
            "FONTE DE RECURSO",
            "ACAO (ID)",
        )
    ):
        linhas = extrair_tabela_de_texto_tsv(texto_copiado, list(COLUNAS_PLANO_ORCAMENTARIO))
        if linhas:
            return linhas

    if any(
        termo in texto_up
        for termo in (
            "GESTAO DE RISCOS",
            "RISCO IDENTIFICADO",
            "ACAO DE MITIGACAO",
            "RESPONSAVEL PELO RISCO",
        )
    ):
        linhas = extrair_tabela_de_texto_tsv(texto_copiado, list(COLUNAS_GESTAO_RISCOS))
        if linhas:
            return linhas

    if any(
        termo in texto_up
        for termo in (
            "TREINAMENTOS TECNICOS",
            "RETENCAO DE TALENTOS",
            "ACOES DE RETENCAO",
        )
    ):
        linhas = extrair_tabela_de_texto_tsv(texto_copiado, list(COLUNAS_CAPACITACAO_5_1))
        if linhas:
            return linhas

    if any(
        termo in texto_up
        for termo in (
            "ACOES MACRO",
            "ACOES MICRO",
            "ACOES EXTRAS",
            "ACOES ESPECIFICAS",
        )
    ):
        linhas = extrair_tabela_de_texto_tsv(texto_copiado, list(COLUNAS_ACOES_ESPECIFICAS_5_3))
        if linhas:
            return linhas

    if any(
        termo in texto_up
        for termo in (
            "INVENTARIO DE NECESSIDADES",
            "DESCRICAO SUSCINTA",
            "DESCRICAO SUCINTA",
            "AREA DEMANDANTE",
            "CLASSIFICACAO DE PRIORIDADE",
        )
    ):
        linhas = extrair_tabela_de_texto_tsv(texto_copiado, list(COLUNAS_INVENTARIO_NECESSIDADES))
        if linhas:
            return linhas

    campos_det = mapeamento.get("campos_detalhados") or {}
    nomes_swot = list(campos_det.keys()) if campos_det else list(COLUNAS_PADRAO_SWOT)
    return extrair_tabela_de_texto_tsv(texto_copiado, nomes_swot)


def copiar_tabela_da_planilha_automaticamente(
    mapeamento: Dict[str, Any],
    coord_planilha_override: Optional[Tuple[int, int]] = None,
) -> List[Dict[str, Any]]:
    """
    Vai até a janela da planilha aberta, clica dentro da tabela,
    executa Ctrl+A e Ctrl+C automaticamente e interpreta todas as colunas e linhas.
    """
    params = obter_parametros_calibrados(mapeamento)
    ponto_planilha = list(coord_planilha_override) if coord_planilha_override else params["planilha"]

    px, py = int(ponto_planilha[0]), int(ponto_planilha[1])
    print(f"📋 Selecionando tabela na planilha aberta em ({px}, {py}) via Ctrl+A e Ctrl+C...")

    copiar_para_clipboard("")
    mover_e_clicar(x=px, y=py, cliques=1, botao="left")
    time.sleep(0.25)

    pressionar_tecla_unica("esc")
    time.sleep(0.08)

    atalho_teclado("ctrl", "a")
    time.sleep(0.22)
    atalho_teclado("ctrl", "c")
    time.sleep(0.35)

    texto_copiado = ler_do_clipboard()
    linhas = _extrair_tabela_auto_detectando_bloco(texto_copiado, mapeamento)

    if not linhas:
        pressionar_tecla_unica("esc")
        time.sleep(0.08)
        pressionar_tecla_unica("home")
        time.sleep(0.12)
        atalho_teclado("ctrl", "a")
        time.sleep(0.25)
        atalho_teclado("ctrl", "c")
        time.sleep(0.40)
        texto_copiado = ler_do_clipboard()
        linhas = _extrair_tabela_auto_detectando_bloco(texto_copiado, mapeamento)

    return linhas


def _determinar_colunas_para_preencher(
    linhas: List[Dict[str, Any]],
    campos_detalhados: Dict[str, Any],
) -> List[str]:
    """
    Determina de forma dinâmica quais colunas da tabela copiada serão preenchidas no sistema,
    usando os cabeçalhos reais da planilha + HTML + mapeamento.
    """
    if not linhas:
        return list(campos_detalhados.keys()) or list(COLUNAS_PADRAO_SWOT)

    colunas_planilha = [
        str(k).strip()
        for k in linhas[0].keys()
        if str(k).strip() and not str(k).startswith("COLUNA_") and not str(k).startswith("_")
    ]

    if not colunas_planilha:
        return list(campos_detalhados.keys()) or list(COLUNAS_PADRAO_SWOT)

    mapa_html = obter_mapa_html_campos()
    mapa_calibrado_norm: Dict[str, str] = {}
    for nome_cal in campos_detalhados.keys():
        n = remover_acentos(nome_cal)
        mapa_calibrado_norm[n] = nome_cal
        if n.endswith("S"):
            mapa_calibrado_norm[n[:-1]] = nome_cal

    colunas_finais: List[str] = []
    for col in colunas_planilha:
        c_norm = remover_acentos(col)
        c_sing = c_norm[:-1] if c_norm.endswith("S") else c_norm
        if c_norm in mapa_calibrado_norm:
            nome_alvo = mapa_calibrado_norm[c_norm]
        elif c_sing in mapa_calibrado_norm:
            nome_alvo = mapa_calibrado_norm[c_sing]
        elif c_norm in mapa_html:
            nome_alvo = str(mapa_html[c_norm]["titulo_exato"]).upper()
        elif c_sing in mapa_html:
            nome_alvo = str(mapa_html[c_sing]["titulo_exato"]).upper()
        else:
            nome_alvo = normalizar(col)

        if nome_alvo not in colunas_finais:
            colunas_finais.append(nome_alvo)

    return colunas_finais


def extrair_colunas_da_tabela(
    linhas: List[Dict[str, Any]],
    campos_detalhados: Optional[Dict[str, Any]] = None,
) -> List[Tuple[str, List[Tuple[int, str]]]]:
    """
    Agrupa os dados da planilha coluna por coluna na sequência exata das colunas:
    Retorna lista de (nome_coluna, [(numero_linha, valor), ...]).
    """
    if campos_detalhados is None:
        mapeamento = carregar_mapeamento()
        campos_detalhados = mapeamento.get("campos_detalhados", {})

    colunas_alvo = _determinar_colunas_para_preencher(linhas, campos_detalhados)
    resultado: List[Tuple[str, List[Tuple[int, str]]]] = []

    for nome_campo in colunas_alvo:
        itens_coluna: List[Tuple[int, str]] = []
        for idx_linha, linha in enumerate(linhas, start=1):
            val = buscar_valor_da_coluna(linha, nome_campo)
            if val:
                itens_coluna.append((idx_linha, val))
        resultado.append((nome_campo, itens_coluna))

    return resultado


def _obter_zonas_proibidas_da_pagina(
    nomes_campos_atuais: List[str],
    campos_detalhados: Dict[str, Any],
    caixas_bloqueadas: List[Dict[str, Any]],
    coord_atual_ref: Optional[Tuple[int, int]] = None,
) -> List[Tuple[int, int]]:
    """
    Reúne apenas as coordenadas (X, Y) de caixas explicitamente marcadas como bloqueadas/protegidas
    (`tipo == "bloqueada"`). Não inclui coordenadas iniciais estáticas de outros campos repetidores
    (ex: Oportunidades/Ameaças/Forças/Fraquezas), pois os cards se expandem e a página rola,
    e a barreira entre blocos vizinhos é calculada dinamicamente em tempo real pelo botão azul do próximo card.
    """
    zonas: List[Tuple[int, int]] = []
    chaves_atuais = {remover_acentos(n) for n in nomes_campos_atuais}

    for item_bloq in caixas_bloqueadas or []:
        if isinstance(item_bloq, dict) and item_bloq.get("barreira_bloco_vizinho"):
            coord_b = item_bloq.get("coordenada")
            if isinstance(coord_b, (list, tuple)) and len(coord_b) >= 2:
                zonas.append((int(coord_b[0]), int(coord_b[1])))

    for outro_nome, info_outro in (campos_detalhados or {}).items():
        if remover_acentos(outro_nome) in chaves_atuais:
            continue
        if not isinstance(info_outro, dict):
            continue
        if str(info_outro.get("tipo", "")).lower() != "bloqueada":
            continue
        coord_o = info_outro.get("sistema_campo_linha1")
        if isinstance(coord_o, (list, tuple)) and len(coord_o) >= 2:
            zonas.append((int(coord_o[0]), int(coord_o[1])))

    return zonas


def _preencher_ou_selecionar_valor_unico(
    nome_campo: str,
    valor: str,
    coord_alvo: Tuple[int, int],
    info_calibrada: Dict[str, Any],
    mapa_html: Dict[str, Any],
    num_linha_planilha: int,
    idx_caixa: int,
    total_caixas: int,
    coord_btn_add: Optional[Tuple[int, int]] = None,
    zonas_proibidas: Optional[List[Tuple[int, int]]] = None,
    y_limite_bloco_vizinho: Optional[int] = None,
    permitir_avancar_abaixo: bool = False,
) -> bool:
    """
    Executa a ação em 1 caixa específica (seja caixa de texto com Ctrl+V ou barra suspensa/toggle
    selecionando a opção vinda da planilha), verificando antes:
    1. Se a coordenada não invade um bloco vizinho/caixa bloqueada (`y_limite_bloco_vizinho` / `zonas_proibidas`).
    2. Se a caixa de texto está VAZIA (`ajustar_coordenada_para_caixa_vazia`), centralizando no `<input>` branco.
    """
    tipo_campo = str(info_calibrada.get("tipo", "caixa_texto")).lower()
    chave_norm = remover_acentos(nome_campo)
    info_h = mapa_html.get(chave_norm)
    if info_h and info_h.get("tag") == "select" and tipo_campo == "caixa_texto":
        tipo_campo = "barra_suspensa"

    if y_limite_bloco_vizinho is not None and int(y_limite_bloco_vizinho) > 220 and int(coord_alvo[1]) >= int(y_limite_bloco_vizinho):
        print(
            f"  🛑 [BLOQUEIO DE BLOCO VIZINHO] Coordenada {coord_alvo} de '{nome_campo}' "
            f"ultrapassou o limite Y={y_limite_bloco_vizinho} do bloco vizinho! Escrita bloqueada com segurança."
        )
        return True

    if tipo_campo == "barra_suspensa":
        opcoes_lista = info_calibrada.get("opcoes_barra_suspensa") or ["Alta", "Média", "Baixa"]
        opcoes_coords = info_calibrada.get("opcoes_coordenadas")
        print(
            f"  🔽 [{nome_campo} — Item {idx_caixa}/{total_caixas} (Linha {num_linha_planilha})] "
            f"Selecionando no toggle/barra suspensa {coord_alvo} -> '{valor}'"
        )
        selecionar_opcao_barra_suspensa(
            coord_caixa=coord_alvo,
            valor_opcao=valor,
            opcoes_barra_suspensa=opcoes_lista,
            opcoes_coordenadas=opcoes_coords,
        )
    else:
        coord_ajustada, esta_vazia, bloqueado_vizinho = ajustar_coordenada_para_caixa_vazia(
            coord_alvo=coord_alvo,
            coord_btn_add=coord_btn_add,
            zonas_proibidas=zonas_proibidas,
            y_limite_bloco_vizinho=y_limite_bloco_vizinho,
            permitir_avancar_abaixo=permitir_avancar_abaixo,
            nome_campo=nome_campo,
        )
        if bloqueado_vizinho:
            print(
                f"  🛑 [BLOQUEIO DE BLOCO VIZINHO] Coordenada {coord_alvo} de '{nome_campo}' "
                "atingiu a barreira do bloco vizinho / caixa bloqueada! Protegendo para NÃO escrever no bloco vizinho."
            )
            return True

        if abs(int(coord_ajustada[1]) - int(coord_alvo[1])) > 25 and esta_vazia and permitir_avancar_abaixo:
            print(
                f"  🔍 [SÓ CAIXA VAZIA] Caixa anterior em {coord_alvo} já ocupada; "
                f"avançando automaticamente para a caixa vazia em {coord_ajustada}!"
            )
            coord_alvo = coord_ajustada
        elif not permitir_avancar_abaixo:
            # Em blocos multi-campos (ex: Metas e ações), preserva o Y unificado da linha horizontal
            # se o ajuste individual tentar desviar mais de 6px dos campos vizinhos da mesma linha!
            if abs(int(coord_ajustada[1]) - int(coord_alvo[1])) <= 6:
                coord_alvo = coord_ajustada
            else:
                coord_alvo = (int(coord_ajustada[0]), int(coord_alvo[1]))
        else:
            coord_alvo = coord_ajustada

        if not esta_vazia and permitir_avancar_abaixo:
            print(
                f"  🛡️ [PROTEÇÃO DE CAMPO PREENCHIDO] A caixa '{nome_campo}' em {coord_alvo} já possui conteúdo! "
                "Preservando valor existente sem sobrescrever."
            )
            return True

        print(
            f"  ✍️  [{nome_campo} — Item {idx_caixa}/{total_caixas} (Linha {num_linha_planilha})] "
            f"Colando na caixa #{idx_caixa} {coord_alvo} -> '{valor}'"
        )
        mover_e_clicar(
            x=coord_alvo[0],
            y=coord_alvo[1],
            cliques=1,
            botao="left",
        )
        time.sleep(max(0.15, ESPERA_CURTA))
        if verificar_aborto():
            return False
        _colar_texto_no_campo(valor)
        time.sleep(ESPERA_CURTA)

    registrar_linha_concluida(
        indice_linha=num_linha_planilha,
        dados_linha={
            "campo": nome_campo,
            "valor": valor,
            "caixa": idx_caixa,
            "coordenada": str(coord_alvo),
        },
    )
    return True


def _executar_passo_grupo_codependente(
    passo: Dict[str, Any],
    linhas: List[Dict[str, Any]],
    campos_detalhados: Dict[str, Any],
    caixas_bloqueadas_pagina: List[Dict[str, Any]],
    mapa_html: Dict[str, Any],
    xn: int,
    ja_rolou_para_baixo: bool,
) -> Tuple[bool, bool]:
    """
    Executa um passo do tipo `grupo_codependente` (ex: Bloco 2 'Inventário de necessidades' na página '3. Diagnóstico'):
    - Para cada linha válida da planilha:
      1. 1º Campo: 'DESCRIÇÃO' (coluna 'DESCRIÇÃO SUSCINTA DE NECESSIDADE PRIORIZADA NO CICLO DO PDTIC')
      2. 2º Campo: 'PRIORIDADE' [toggle] (coluna 'CLASSIFICAÇÃO DE PRIORIDADE')
      3. 3º Campo: 'ÁREA DEMANDANTE' (coluna 'ÁREA DEMANDANTE' — valor selecionado da lista suspensa)
      4. 4º Campo: 'JUSTIFICATIVA' (coluna 'JUSTIFICATIVA' — campo maior abaixo dos 3 primeiros, com scroll extra para baixo)
      5. Somente após completar os 4 campos da mesma linha, clica em
         '<button type="button" class="btn btn-sm btn-outline-primary" data-repeater-add="needs">Adicionar necessidade</button>',
         aguarda `ESPERA_APOS_ADICIONAR` e avança para a próxima linha.
    Retorna `(sucesso, ja_rolou_para_baixo)`.
    """
    nome_grupo = str(passo.get("nome_passo") or "Inventário de necessidades")
    campos_grupo = [str(c) for c in (passo.get("campos") or []) if str(c).strip()]
    if not campos_grupo:
        return True, ja_rolou_para_baixo

    # Coleta as linhas da planilha que possuem conteúdo real para os campos deste grupo
    # (Ignora linhas vazias do template que só tenham o toggle 'CLASSIFICAÇÃO DE PRIORIDADE' pré-selecionado)
    linhas_grupo: List[Tuple[int, Dict[str, str]]] = []
    for idx_linha, linha in enumerate(linhas, start=1):
        valores_linha: Dict[str, str] = {}
        tem_conteudo_nao_toggle = False
        for c_nome in campos_grupo:
            val = buscar_valor_da_coluna(linha, c_nome)
            info_c_tmp = _buscar_info_campo_calibrado(c_nome, campos_detalhados)
            eh_toggle = (
                info_c_tmp.get("tipo") == "barra_suspensa"
                or _resolver_chave_sinonimo(c_nome) == "PRIORIDADE"
            )
            if val:
                valores_linha[c_nome] = val
                if not eh_toggle:
                    tem_conteudo_nao_toggle = True
            elif eh_toggle:
                opcoes_padrao = info_c_tmp.get("opcoes") or ["Alta", "Média", "Baixa"]
                valores_linha[c_nome] = str(opcoes_padrao[0])

        if valores_linha and (tem_conteudo_nao_toggle or len(campos_grupo) == 1):
            if linha.get("_ORDEM_SEQUENCIAL") is not None:
                valores_linha["_ORDEM_SEQUENCIAL"] = str(linha["_ORDEM_SEQUENCIAL"])
            num_l_exibicao = int(linha.get("_IDX_LINHA_ORIGINAL") or idx_linha)
            linhas_grupo.append((num_l_exibicao, valores_linha))

    if not linhas_grupo:
        print(
            f"\n⚪ Grupo Codependente '{nome_grupo}' ({' -> '.join(campos_grupo)}): "
            "nenhuma linha preenchida encontrada na planilha."
        )
        return True, ja_rolou_para_baixo

    if LIMITE_LINHAS_TESTE is not None and LIMITE_LINHAS_TESTE > 0 and len(linhas_grupo) > LIMITE_LINHAS_TESTE:
        total_orig_grupo = len(linhas_grupo)
        linhas_grupo = linhas_grupo[:LIMITE_LINHAS_TESTE]
        print(
            f"  🔢 [LIMITE DE TESTE ATIVO] Copiando/colando apenas as primeiras "
            f"{len(linhas_grupo)} de {total_orig_grupo} linha(s) de '{nome_grupo}'!"
        )

    # Obtém as coordenadas iniciais salvas de cada campo do grupo e do botão 'Adicionar necessidade'
    coords_iniciais: List[Tuple[int, int]] = []
    infos_campos: List[Dict[str, Any]] = []
    coord_btn_grupo: Optional[Tuple[int, int]] = None
    passo_y_grupo = passo.get("passo_y_grupo")

    if isinstance(passo.get("botao_adicionar"), (list, tuple)) and len(passo["botao_adicionar"]) >= 2:
        coord_btn_grupo = (
            max(650, min(1860, int(passo["botao_adicionar"][0]))),
            int(passo["botao_adicionar"][1]),
        )

    for c_nome in campos_grupo:
        info_c = _buscar_info_campo_calibrado(c_nome, campos_detalhados)
        infos_campos.append(info_c)
        c_coord = info_c.get("sistema_campo_linha1")
        if isinstance(c_coord, (list, tuple)) and len(c_coord) >= 2:
            coords_iniciais.append(
                (max(650, min(1760, int(c_coord[0]))), int(c_coord[1]))
            )
        else:
            coords_iniciais.append((910, 605))
        if coord_btn_grupo is None and isinstance(info_c.get("botao_adicionar"), (list, tuple)):
            coord_btn_grupo = (
                max(650, min(1860, int(info_c["botao_adicionar"][0]))),
                int(info_c["botao_adicionar"][1]),
            )
        if not passo_y_grupo and info_c.get("passo_y_grupo"):
            passo_y_grupo = int(info_c["passo_y_grupo"])

    # Localiza o cabeçalho do bloco (ex: 'Inventário de necessidades' no final da página 3,
    # ou 'Metas e ações' / 'Plano orçamentário' na página 4)
    termo_card_busca = str(passo.get("titulo_card_busca") or nome_grupo)
    coord_c1_salva = coords_iniciais[0]
    eh_bloco_metas = (
        "METAS" in remover_acentos(nome_grupo)
        or "METAS" in remover_acentos(termo_card_busca)
    )
    eh_bloco_orcamentario = (
        "ORCAMENTARIO" in remover_acentos(nome_grupo)
        or "ORCAMENTARIO" in remover_acentos(termo_card_busca)
    )
    eh_bloco_riscos = (
        "RISCO" in remover_acentos(nome_grupo)
        or "RISCO" in remover_acentos(termo_card_busca)
    )
    eh_bloco_kpis = (
        "KPI" in remover_acentos(nome_grupo)
        or "INDICADORES DE SUCESSO" in remover_acentos(nome_grupo)
        or "KPI" in remover_acentos(termo_card_busca)
    )
    ancorado_no_topo = bool(
        passo.get("ancorado_no_topo", False)
        or eh_bloco_metas
        or eh_bloco_orcamentario
        or eh_bloco_riscos
        or eh_bloco_kpis
    )
    y_limite_bloco_vizinho = passo.get("y_limite_bloco_vizinho")
    aprendizado_3_linhas = passo.get("aprendizado_3_linhas")

    if eh_bloco_metas:
        # Como 'Metas e ações' já começa no topo da página 4 (ir_para_topo_no_inicio=True),
        # usa diretamente as coordenadas calibradas no topo sem deslocar a página com Ctrl+F
        coord_btn_localizado = coord_btn_grupo
        delta_y_inicial = 0
    else:
        _, coord_btn_localizado, ja_rolou_para_baixo = localizar_card_com_ctrl_f_e_visao(
            nome_campo=termo_card_busca,
            coord_campo_salva=coord_c1_salva,
            coord_btn_salva=coord_btn_grupo,
            eh_quadrante_inferior=(not (eh_bloco_orcamentario or eh_bloco_riscos or eh_bloco_kpis)),
            ja_rolou_para_baixo=ja_rolou_para_baixo,
            incluir_botoes_largos=True,
            posicionar_no_topo=(eh_bloco_orcamentario or eh_bloco_riscos or eh_bloco_kpis),
        )

        if coord_btn_localizado is not None and coord_btn_grupo is not None and 205 <= int(coord_btn_localizado[1]) <= 860:
            delta_y_inicial = int(coord_btn_localizado[1]) - int(coord_btn_grupo[1])
            if not (eh_bloco_orcamentario or eh_bloco_riscos or eh_bloco_kpis) and abs(delta_y_inicial) > 260:
                delta_y_inicial = 0
        else:
            delta_y_inicial = 0

    coords_atuais: List[Tuple[int, int]] = [
        (max(650, min(1760, int(cx))), max(218, int(cy + delta_y_inicial)))
        for cx, cy in coords_iniciais
    ]
    if coord_btn_localizado is not None:
        coord_btn_atual: Optional[Tuple[int, int]] = coord_btn_localizado
    elif coord_btn_grupo is not None:
        coord_btn_atual = (int(coord_btn_grupo[0]), int(coord_btn_grupo[1] + delta_y_inicial))
    else:
        coord_btn_atual = None

    zonas_proibidas = _obter_zonas_proibidas_da_pagina(
        nomes_campos_atuais=campos_grupo,
        campos_detalhados=campos_detalhados,
        caixas_bloqueadas=caixas_bloqueadas_pagina,
        coord_atual_ref=coords_atuais[0],
    )
    if y_limite_bloco_vizinho is None:
        for item_bloq in caixas_bloqueadas_pagina or []:
            if isinstance(item_bloq, dict) and item_bloq.get("barreira_bloco_vizinho"):
                coord_b = item_bloq.get("coordenada")
                if isinstance(coord_b, (list, tuple)) and len(coord_b) >= 2:
                    y_limite_bloco_vizinho = int(coord_b[1]) - 18
                    break

    resumo_tipos = []
    for idx_res, (c_nome, info_c) in enumerate(zip(campos_grupo, infos_campos), start=1):
        t_c = "TOGGLE/SELECT" if info_c.get("tipo") == "barra_suspensa" else "TEXTO"
        resumo_tipos.append(f"{idx_res}º {c_nome} [{t_c}]")

    nome_botao_html = str(
        passo.get("texto_botao_adicionar")
        or (
            "Adicionar ação"
            if eh_bloco_metas
            else (
                "Adicionar linha"
                if eh_bloco_orcamentario
                else ("Adicionar risco" if eh_bloco_riscos else "Adicionar necessidade")
            )
        )
    )
    print(
        f"\n🔗 BLOCO CODEPENDENTE '{nome_grupo}' ({len(campos_grupo)} CAMPOS NA SEQUÊNCIA): "
        f"{' -> '.join(resumo_tipos)} -> Botão '{nome_botao_html}' ({len(linhas_grupo)} linha(s) na planilha)"
    )

    coords_base_fundo = list(coords_atuais)
    ticks_rolados = 0

    # Verifica se a 1ª linha visível na tela já está preenchida (para NUNCA sobrescrever caixas que já tenham texto!)
    if coords_atuais and coord_btn_atual is not None:
        img_ini, off_ini = capturar_tela_sistema(coords_atuais[0][0])
        if img_ini is not None:
            if ancorado_no_topo:
                if eh_bloco_orcamentario or eh_bloco_riscos or eh_bloco_kpis:
                    ys_x_ini = detectar_botoes_remover_da_coluna(
                        img_ini,
                        off_ini,
                        x_btn_add=int(coord_btn_atual[0]),
                        y_btn_add=int(coord_btn_atual[1]) + 20,
                        y_limite_inferior=920,
                    )
                    if ys_x_ini:
                        y_x1 = ys_x_ini[0]
                        y_prim_base = int(coords_iniciais[0][1])
                        coords_atuais = [
                            (
                                cx,
                                centralizar_y_na_caixa_branca(
                                    img_ini, off_ini, cx, y_x1 + int(cy_orig - y_prim_base)
                                ),
                            )
                            for (cx, _), (_, cy_orig) in zip(coords_atuais, coords_iniciais)
                        ]
                    else:
                        coords_atuais = [
                            (cx, centralizar_y_na_caixa_branca(img_ini, off_ini, cx, cy))
                            for cx, cy in coords_atuais
                        ]
                else:
                    coords_atuais = [
                        (cx, centralizar_y_na_caixa_branca(img_ini, off_ini, cx, cy))
                        for cx, cy in coords_atuais
                    ]
                coords_base_fundo = list(coords_atuais)
            vazia_c0 = eh_caixa_de_texto_vazia(
                img_ini, off_ini, int(coords_atuais[0][0]), int(coords_atuais[0][1]), None, nome_campo=campos_grupo[0]
            )
            vazia_cn = eh_caixa_de_texto_vazia(
                img_ini, off_ini, int(coords_atuais[-1][0]), int(coords_atuais[-1][1]), None, nome_campo=campos_grupo[-1]
            )
            if not vazia_c0 and not vazia_cn:
                print(
                    f"  🔍 [SÓ CAIXA VAZIA] A 1ª linha de '{nome_grupo}' já possui dados preenchidos; "
                    f"clicando em '{nome_botao_html}' para abrir uma nova linha VAZIA..."
                )
                coords_atuais, coord_btn_atual, ticks_ini = clicar_adicionar_e_avancar_grupo_codependente(
                    coords_campos_atuais=coords_atuais,
                    coord_btn_atual=coord_btn_atual,
                    passo_y_grupo=int(passo_y_grupo) if passo_y_grupo else None,
                    zonas_proibidas=zonas_proibidas,
                    forcar_scroll_padrao=True,
                    indice_caixa=1,
                    coords_base_fundo=coords_base_fundo,
                    ancorado_no_topo=ancorado_no_topo,
                    texto_botao_adicionar=nome_botao_html,
                    y_limite_bloco_vizinho=y_limite_bloco_vizinho,
                    aprendizado_3_linhas=aprendizado_3_linhas,
                )
                ticks_rolados += ticks_ini

    for idx_bloco, (num_linha_planilha, valores_linha) in enumerate(linhas_grupo):
        if verificar_aborto():
            return False, ja_rolou_para_baixo

        rotulo_item = "Ação" if ancorado_no_topo else "Necessidade"
        info_ordem_seq = (
            f"ORDEM SEQUENCIAL: {valores_linha['_ORDEM_SEQUENCIAL']} | "
            if valores_linha.get("_ORDEM_SEQUENCIAL")
            else ""
        )
        print(
            f"\n  🧩 [{rotulo_item} {idx_bloco + 1}/{len(linhas_grupo)} ({info_ordem_seq}Linha {num_linha_planilha} da planilha)] "
            f"Preenchendo os {len(campos_grupo)} campos na sequência antes de clicar em '{nome_botao_html}'..."
        )

        # AUTOCORREÇÃO, PROTEÇÃO DE CAMPO PREENCHIDO E ALINHAMENTO DE LINHA HORIZONTAL (a partir da 2ª caixa):
        # 1. Garante que todos os campos na mesma linha horizontal do sub-card (ex: ID, OBJETIVO ESTRATÉGICO e META)
        #    compartilhem exatamente o mesmo Y centralizado nas bordas reais do <input>.
        # 2. Confere se o sub-card mirado já possui texto preenchido; se estiver ocupado, NUNCA sobrescreve:
        #    avança automaticamente +ΔY (passo_y_grupo) para o sub-card vazio abaixo!
        if idx_bloco >= 1 and coords_atuais:
            img_pre, off_pre = capturar_tela_sistema(coords_atuais[0][0])
            if img_pre is not None:
                coords_cent = [
                    (cx, centralizar_y_na_caixa_branca(img_pre, off_pre, cx, cy))
                    for cx, cy in coords_atuais
                ]
                coords_unificadas: List[Tuple[int, int]] = []
                for i_c, (cx_i, cy_i) in enumerate(coords_cent):
                    ys_linha_h = [
                        coords_cent[j_c][1]
                        for j_c in range(len(coords_cent))
                        if abs(coords_iniciais[j_c][1] - coords_iniciais[i_c][1]) <= 12
                    ]
                    y_med = sorted(ys_linha_h)[len(ys_linha_h) // 2] if ys_linha_h else cy_i
                    coords_unificadas.append((cx_i, int(y_med)))
                coords_atuais = coords_unificadas

                idx_ref_prim = 1 if (len(coords_atuais) >= 2 and abs(coords_iniciais[1][1] - coords_iniciais[0][1]) <= 12) else 0
                v_c0 = eh_caixa_de_texto_vazia(
                    img_pre,
                    off_pre,
                    int(coords_atuais[idx_ref_prim][0]),
                    int(coords_atuais[idx_ref_prim][1]),
                    None,
                    nome_campo=campos_grupo[idx_ref_prim],
                )
                v_cn = eh_caixa_de_texto_vazia(
                    img_pre,
                    off_pre,
                    int(coords_atuais[-1][0]),
                    int(coords_atuais[-1][1]),
                    None,
                    nome_campo=campos_grupo[-1],
                )
                if not v_c0 and not v_cn:
                    passo_corr = int(
                        passo_y_grupo
                        or (188 if (eh_bloco_orcamentario or eh_bloco_riscos) else (110 if eh_bloco_kpis else 286))
                    )
                    print(
                        f"  🛡️ [PROTEÇÃO DE CAMPO PREENCHIDO] Sub-card em Y={coords_atuais[0][1]} já está preenchido! "
                        f"Avançando +{passo_corr}px para o próximo sub-card vazio abaixo..."
                    )
                    if int(coords_atuais[-1][1]) + passo_corr > 865:
                        ticks_c = max(2, min(4, int(round(passo_corr / 100.0))))
                        rolar_pagina(-ticks_c, x=1230, y=500)
                        time.sleep(0.35)
                        img_pre2, off_pre2 = capturar_tela_sistema(coords_atuais[0][0])
                        if img_pre2 is not None:
                            coords_atuais = [
                                (cx, centralizar_y_na_caixa_branca(img_pre2, off_pre2, cx, cy))
                                for cx, cy in coords_atuais
                            ]
                    else:
                        coords_atuais = [
                            (cx, centralizar_y_na_caixa_branca(img_pre, off_pre, cx, int(cy + passo_corr)))
                            for cx, cy in coords_atuais
                        ]

        campos_preenchidos_na_linha = 0
        for idx_c, c_nome in enumerate(campos_grupo):
            if verificar_aborto():
                return False, ja_rolou_para_baixo
            valor_c = valores_linha.get(c_nome, "").strip()
            if not valor_c:
                continue

            info_c = infos_campos[idx_c]
            y_campo_atual = coords_atuais[idx_c][1]
            scroll_antes = int(info_c.get("scroll_antes_campo", 0))

            # Só rola antes se o campo passar de Y = 880!
            if scroll_antes < 0 and y_campo_atual < 880:
                scroll_antes = 0

            if scroll_antes == 0 and y_campo_atual > 890:
                scroll_antes = -max(2, min(5, int(round((y_campo_atual - 720) / 90.0))))

            if scroll_antes != 0:
                print(
                    f"  📜 [Scroll Down -> '{c_nome}'] Rolando {scroll_antes} ticks para baixo para "
                    f"acompanhar a expansão da página e posicionar '{c_nome}'..."
                )
                x_ref_sc = coords_atuais[idx_c][0]
                img_sc_antes, off_sc_antes = capturar_tela_sistema(x_ref_sc)
                rolar_pagina(scroll_antes, x=1230, y=500)
                time.sleep(0.35)
                img_sc_depois, _ = capturar_tela_sistema(x_ref_sc)
                delta_sc = _medir_scroll_por_template(
                    img_sc_antes,
                    img_sc_depois,
                    off_sc_antes,
                    x_ref_sc,
                    passos_scroll=scroll_antes,
                )
                if delta_sc is not None and abs(delta_sc) >= 8:
                    coords_atuais = [
                        (cx, int(cy - delta_sc)) for cx, cy in coords_atuais
                    ]
                    if coord_btn_atual is not None:
                        coord_btn_atual = (
                            coord_btn_atual[0],
                            int(coord_btn_atual[1] - delta_sc),
                        )

            ok = _preencher_ou_selecionar_valor_unico(
                nome_campo=c_nome,
                valor=valor_c,
                coord_alvo=coords_atuais[idx_c],
                info_calibrada=info_c,
                mapa_html=mapa_html,
                num_linha_planilha=num_linha_planilha,
                idx_caixa=idx_bloco + 1,
                total_caixas=len(linhas_grupo),
                coord_btn_add=coord_btn_atual,
                zonas_proibidas=zonas_proibidas,
                y_limite_bloco_vizinho=y_limite_bloco_vizinho,
                permitir_avancar_abaixo=False,
            )
            if not ok:
                return False, ja_rolou_para_baixo

            # AUTOCORREÇÃO PÓS-COLAGEM no 1º campo da caixa:
            # Se após colar no 1º campo a caixa continuou 100% em branco (ex: perda de foco),
            # centraliza novamente na caixa branca e repete o clique + Ctrl+V!
            if idx_c == 0 and info_c.get("tipo") != "barra_suspensa":
                img_pos_c, off_pos_c = capturar_tela_sistema(coords_atuais[0][0])
                if img_pos_c is not None and eh_caixa_de_texto_vazia(
                    img_pos_c,
                    off_pos_c,
                    int(coords_atuais[0][0]),
                    int(coords_atuais[0][1]),
                    None,
                    nome_campo=c_nome,
                ):
                    cy_c0 = centralizar_y_na_caixa_branca(
                        img_pos_c, off_pos_c, int(coords_atuais[0][0]), int(coords_atuais[0][1])
                    )
                    if len(coords_atuais) >= 2 and abs(coords_iniciais[1][1] - coords_iniciais[0][1]) <= 12:
                        cy_c1 = centralizar_y_na_caixa_branca(
                            img_pos_c, off_pos_c, int(coords_atuais[1][0]), int(coords_atuais[1][1])
                        )
                        cy_retry = cy_c1 if abs(cy_c1 - cy_c0) <= 18 else cy_c0
                    else:
                        cy_retry = cy_c0
                    print(
                        f"  🛠️ [AUTOCORREÇÃO PÓS-COLAGEM] Reforçando foco e colagem em '{c_nome}' "
                        f"em ({coords_atuais[0][0]}, {cy_retry})..."
                    )
                    copiar_para_clipboard(valor_c)
                    mover_e_clicar(int(coords_atuais[0][0]), int(cy_retry))
                    time.sleep(0.08)
                    atalho_teclado("ctrl", "a")
                    time.sleep(0.04)
                    atalho_teclado("ctrl", "v")
                    time.sleep(0.12)

            campos_preenchidos_na_linha += 1

        tem_proximo_bloco = (idx_bloco + 1) < len(linhas_grupo)
        if tem_proximo_bloco and coord_btn_atual is not None and campos_preenchidos_na_linha > 0:
            aviso_subir = (
                f" [Subindo a página antes para localizar o botão '{nome_botao_html}' no topo]"
                if (idx_bloco + 1) >= 2
                else ""
            )
            print(
                f"  ➕ [{nome_grupo}] {len(campos_grupo)} campos da caixa {idx_bloco + 1} preenchidos!{aviso_subir} "
                f"Clicando em '{nome_botao_html}', aguardando {ESPERA_APOS_ADICIONAR:.2f}s e posicionando os novos campos..."
            )
            coords_atuais, coord_btn_atual, ticks_passo = (
                clicar_adicionar_e_avancar_grupo_codependente(
                    coords_campos_atuais=coords_atuais,
                    coord_btn_atual=coord_btn_atual,
                    passo_y_grupo=int(passo_y_grupo) if passo_y_grupo else None,
                    zonas_proibidas=zonas_proibidas,
                    forcar_scroll_padrao=True,
                    indice_caixa=idx_bloco + 1,
                    coords_base_fundo=coords_base_fundo,
                    ancorado_no_topo=ancorado_no_topo,
                    texto_botao_adicionar=nome_botao_html,
                    y_limite_bloco_vizinho=y_limite_bloco_vizinho,
                    aprendizado_3_linhas=aprendizado_3_linhas,
                )
            )
            ticks_rolados += ticks_passo

    print(
        f"\n  ✅ Bloco '{nome_grupo}' 100% concluído "
        f"({len(linhas_grupo)}/{len(linhas_grupo)} linha(s) preenchida(s))!"
    )
    return True, ja_rolou_para_baixo


def _montar_roteiro_passos_da_pagina(
    padrao_pagina: Dict[str, Any],
    colunas_com_itens: List[Tuple[str, List[Tuple[int, str]]]],
    campos_detalhados: Dict[str, Any],
) -> List[Dict[str, Any]]:
    """
    Combina os `passos_pagina` salvos no treinamento da página com as colunas presentes na planilha:
    - Quando `passos_pagina` já está definido no bloco, respeita estritamente os passos ensinados
      (impedindo que colunas extras da planilha vazem para um bloco vizinho na mesma página!).
    - Apenas quando `passos_pagina` está vazio monta um roteiro automático pelas colunas da planilha.
    """
    passos_salvos = padrao_pagina.get("passos_pagina") or []
    roteiro: List[Dict[str, Any]] = []
    campos_cobertos = set()

    for p in passos_salvos:
        if not isinstance(p, dict):
            continue
        lista_c = [str(c) for c in (p.get("campos") or []) if str(c).strip()]
        if not lista_c:
            continue
        roteiro.append(dict(p))
        for c_nome in lista_c:
            campos_cobertos.add(remover_acentos(c_nome))
            campos_cobertos.add(_resolver_chave_sinonimo(c_nome))

    if roteiro:
        return roteiro

    for nome_col, _ in colunas_com_itens:
        chave_col = remover_acentos(nome_col)
        canon_col = _resolver_chave_sinonimo(nome_col)
        if chave_col in campos_cobertos or canon_col in campos_cobertos:
            continue
        info_c = _buscar_info_campo_calibrado(nome_col, campos_detalhados)
        tem_add = bool(info_c.get("tem_adicionar", bool(info_c.get("botao_adicionar", True))))
        roteiro.append(
            {
                "tipo_passo": "coluna_independente" if tem_add else "campo_unico",
                "nome_passo": nome_col,
                "campos": [nome_col],
            }
        )
        campos_cobertos.add(chave_col)
        campos_cobertos.add(canon_col)

    return roteiro


def _copiar_tabela_da_guia_atual_na_planilha(
    coord_celula_planilha: Tuple[int, int],
    nomes_esperados: List[str],
    nome_guia: str,
) -> List[Dict[str, Any]]:
    """
    Executa a cópia segura (`ESC -> Ctrl+A -> Ctrl+C`) na guia que já está aberta na planilha
    e extrai as linhas da tabela.
    """
    cx, cy = int(coord_celula_planilha[0]), int(coord_celula_planilha[1])
    print(
        f"📋 Copiando dados da guia '{nome_guia}' em ({cx}, {cy}) "
        "com proteção de toggles/listas (ESC -> Ctrl+A -> Ctrl+C)..."
    )
    copiar_para_clipboard("")
    mover_e_clicar(x=cx, y=cy, cliques=1, botao="left")
    time.sleep(0.28)

    pressionar_tecla_unica("esc")
    time.sleep(0.10)

    atalho_teclado("ctrl", "a")
    time.sleep(0.25)
    atalho_teclado("ctrl", "c")
    time.sleep(0.40)

    texto_copiado = ler_do_clipboard()
    linhas = extrair_tabela_de_texto_tsv(texto_copiado, nomes_esperados)

    if not linhas:
        # Se a planilha estava rolada para a direita, 'Home' volta para a coluna A da mesma linha
        pressionar_tecla_unica("esc")
        time.sleep(0.08)
        pressionar_tecla_unica("home")
        time.sleep(0.15)
        atalho_teclado("ctrl", "a")
        time.sleep(0.25)
        atalho_teclado("ctrl", "c")
        time.sleep(0.45)
        texto_copiado = ler_do_clipboard()
        linhas = extrair_tabela_de_texto_tsv(texto_copiado, nomes_esperados)

    if not linhas:
        # Fallback adicional caso a célula clicada estivesse fora da tabela
        atalho_teclado("ctrl", "home")
        time.sleep(0.15)
        pressionar_tecla_unica("down", repeticoes=4, intervalo=0.04)
        pressionar_tecla_unica("right", repeticoes=1, intervalo=0.04)
        time.sleep(0.10)
        atalho_teclado("ctrl", "a")
        time.sleep(0.25)
        atalho_teclado("ctrl", "c")
        time.sleep(0.45)
        texto_copiado = ler_do_clipboard()
        linhas = extrair_tabela_de_texto_tsv(texto_copiado, nomes_esperados)

    if not linhas:
        # Fallback Ctrl+A duplo para incluir cabeçalhos mesclados (ex: PRAZO -> QUADRANTE / ANO)
        pressionar_tecla_unica("esc")
        time.sleep(0.08)
        atalho_teclado("ctrl", "a")
        time.sleep(0.15)
        atalho_teclado("ctrl", "a")
        time.sleep(0.25)
        atalho_teclado("ctrl", "c")
        time.sleep(0.45)
        texto_copiado = ler_do_clipboard()
        linhas = extrair_tabela_de_texto_tsv(texto_copiado, nomes_esperados)

    if linhas:
        cols_lidas = [k for k in linhas[0].keys() if not str(k).startswith("_")]
        amostra_1 = {
            k: (str(v)[:35] + "..." if len(str(v)) > 35 else str(v))
            for k, v in linhas[0].items()
            if not str(k).startswith("_")
        }
        print(f"   📌 Colunas identificadas na guia '{nome_guia}': {cols_lidas}")
        if any(l.get("_ORDEM_SEQUENCIAL") is not None for l in linhas):
            ordens_seq = [l.get("_ORDEM_SEQUENCIAL") for l in linhas if l.get("_ORDEM_SEQUENCIAL") is not None]
            print(f"   🔢 Linhas ordenadas por 'ORDEM SEQUENCIAL' crescente: {ordens_seq}")
        print(f"   🔎 Amostra 1ª Linha a colar: {amostra_1}")
    print(f"✅ Guia '{nome_guia}' lida com sucesso: {len(linhas)} linha(s) válida(s) encontrada(s)!")
    print("-" * 72)
    return linhas


def trocar_guia_planilha_e_copiar_tabela(
    coord_guia_planilha: Optional[Tuple[int, int]],
    coord_celula_planilha: Tuple[int, int],
    nomes_esperados: List[str],
    nome_guia: str = "3.2. Inventário de necessidades",
    rotulo_ligacao: str = "TROCA DE GUIA NA PLANILHA",
    coord_toggle_todas_guias: Optional[Tuple[int, int]] = None,
    coord_item_no_toggle: Optional[Tuple[int, int]] = None,
) -> List[Dict[str, Any]]:
    """
    Seleciona a guia da planilha (preferencialmente via Toggle de Todas as Guias '≡' + Nome de Referência,
    sem depender de posição variável na barra inferior) e só depois copia a tabela:
    1. Se `coord_toggle_todas_guias` estiver definido, clica no Toggle '≡' e escolhe `nome_guia` pelo nome
       (via OCR nativo do Windows ou `coord_item_no_toggle`). Caso contrário, usa `coord_guia_planilha`.
    2. Clica dentro da tabela da guia (`coord_celula_planilha`), pressiona ESC e executa Ctrl+A e Ctrl+C.
    3. Extrai todas as colunas da tabela e retorna as linhas prontas para preencher no sistema.
    """
    print("\n" + "-" * 72)
    print(f"🔗 [{rotulo_ligacao}] Selecionando a guia '{nome_guia}' na planilha...")
    selecionar_guia_planilha_via_toggle_e_nome(
        nome_guia=nome_guia,
        coord_toggle_todas_guias=coord_toggle_todas_guias,
        coord_item_no_toggle=coord_item_no_toggle,
        coord_guia_fallback=coord_guia_planilha,
    )
    return _copiar_tabela_da_guia_atual_na_planilha(
        coord_celula_planilha=coord_celula_planilha,
        nomes_esperados=nomes_esperados,
        nome_guia=nome_guia,
    )


def _confirmar_correspondencia_guia_planilha(
    linhas: List[Dict[str, Any]],
    nomes_esperados: List[str],
    nome_guia: str,
) -> bool:
    """
    Confirma se as colunas extraídas da planilha correspondem de fato à guia esperada
    ('3.1. Análise SWOT', '3.2. Inventário de necessidades', '4.1. Metas e Ações', etc.)
    antes de iniciar o preenchimento no sistema.
    """
    if not linhas:
        return False

    esperados_norm = [remover_acentos(c) for c in nomes_esperados]
    tem_valor_esperado = any(
        bool(buscar_valor_da_coluna(l, col_esp))
        for l in linhas
        for col_esp in nomes_esperados
    )
    if not tem_valor_esperado:
        return False

    # Confirmação específica para impedir mistura entre SWOT, Inventário, Metas/Orçamento/Riscos, Capacitação e Monitoramento
    cols_lidas_norm = {remover_acentos(k) for k in linhas[0].keys() if not str(k).startswith("_")}
    if any("FORCA" in e for e in esperados_norm):
        return any("FORCA" in c or "FRAQUEZA" in c or "OPORTUNIDADE" in c or "AMEACA" in c for c in cols_lidas_norm)
    if any("AREA DEMANDANTE" in e or "JUSTIFICATIVA" in e for e in esperados_norm):
        return any("DESCRICAO" in c or "PRIORIDADE" in c or "AREA DEMANDANTE" in c for c in cols_lidas_norm)
    if any("DESCRICAO DO CUSTO" in e or "FONTE DE RECURSO" in e or "ACAO (ID)" in e for e in esperados_norm):
        return any("CUSTO" in c or "FONTE" in c or "ACAO (ID)" in c or "ANO 1" in c for c in cols_lidas_norm)
    if any("RISCO IDENTIFICADO" in e or "ACAO DE MITIGACAO" in e or "RESPONSAVEL PELO RISCO" in e for e in esperados_norm):
        return any("RISCO" in c or "MITIGACAO" in c or "PROBABILIDADE" in c or "IMPACTO" in c for c in cols_lidas_norm)
    if any("TREINAMENTOS TECNICOS" in e or "RETENCAO DE TALENTOS" in e for e in esperados_norm):
        return any("TREINAMENTO" in c or "RETENCAO" in c or "TALENTO" in c for c in cols_lidas_norm)
    if any("ACOES MACRO" in e or "ACOES MICRO" in e or "ACOES EXTRAS" in e for e in esperados_norm):
        return any("MACRO" in c or "MICRO" in c or "EXTRA" in c for c in cols_lidas_norm)
    if any("FORMULA" in e for e in esperados_norm) and not any("OBJETIVO" in e for e in esperados_norm):
        return any("INDICADOR" in c or "FORMULA" in c or "METRICA" in c for c in cols_lidas_norm)
    if any("OBJETIVO ESTRATEGICO" in e or "INDICADOR" in e for e in esperados_norm):
        return any("OBJETIVO" in c or "META" in c or "INDICADOR" in c or "RESPONSAVEL" in c for c in cols_lidas_norm)

    print(f"   ✅ Confirmada correspondência da guia da planilha '{nome_guia}'!")
    return True


def _garantir_coord_toggle_ou_guia_planilha(
    padrao_pagina: Dict[str, Any],
    mapeamento: Dict[str, Any],
    nome_guia_padrao: str,
    solicitar_se_ausente: bool = True,
    forcar_recalibracao: bool = False,
) -> Tuple[Optional[Tuple[int, int]], Optional[Tuple[int, int]], Optional[Tuple[int, int]]]:
    """
    Retorna `(coord_toggle_todas_guias, coord_item_no_toggle, coord_guia_planilha)`.
    Prioriza o botão fixo 'Toggle de Todas as Guias' (`≡`) na planilha para que a escolha da aba
    seja feita pelo nome de referência (`nome_guia_planilha`) sem depender da posição horizontal da aba.
    """
    coord_toggle = mapeamento.get("coord_toggle_todas_guias") or padrao_pagina.get("coord_toggle_todas_guias")
    coord_item_toggle = padrao_pagina.get("coord_item_no_toggle")
    coord_guia = padrao_pagina.get("coord_guia_planilha")

    t_tuple = (
        (int(coord_toggle[0]), int(coord_toggle[1]))
        if isinstance(coord_toggle, (list, tuple)) and len(coord_toggle) >= 2 and coord_toggle[0] is not None
        else None
    )
    i_tuple = (
        (int(coord_item_toggle[0]), int(coord_item_toggle[1]))
        if isinstance(coord_item_toggle, (list, tuple)) and len(coord_item_toggle) >= 2 and coord_item_toggle[0] is not None
        else None
    )
    g_tuple = (
        (int(coord_guia[0]), int(coord_guia[1]))
        if isinstance(coord_guia, (list, tuple)) and len(coord_guia) >= 2 and coord_guia[0] is not None
        else None
    )

    if not forcar_recalibracao and (t_tuple is not None or g_tuple is not None):
        return t_tuple, i_tuple, g_tuple

    if not solicitar_se_ausente:
        return t_tuple, i_tuple, g_tuple

    nome_guia = str(padrao_pagina.get("nome_guia_planilha") or nome_guia_padrao)
    print("\n" + "=" * 72)
    print(f"📑 [CALIBRAÇÃO DO TOGGLE DE TODAS AS GUIAS ('≡') + ABA '{nome_guia}']")
    print(
        "👉 Para o robô NÃO depender da posição variável das abas na planilha:\n"
        "   1º Posicione o mouse sobre o botão TOGGLE '≡' (Todas as páginas / Todas as guias)\n"
        "      no canto inferior esquerdo da planilha e aperte [F8]:"
    )
    tx, ty = aguardar_tecla_f8()
    t_tuple = (int(tx), int(ty))
    mapeamento["coord_toggle_todas_guias"] = [int(tx), int(ty)]
    padrao_pagina["coord_toggle_todas_guias"] = [int(tx), int(ty)]
    padrao_pagina["nome_guia_planilha"] = nome_guia
    print(f"   ✅ Toggle de Todas as Guias salvo em: ({tx}, {ty})!")

    # Clica automaticamente no Toggle para abrir a lista e tentar OCR ou permitir apontar o item
    print(f"   🔄 Abrindo o Toggle '≡' em ({tx}, {ty}) para localizar '{nome_guia}' pelo nome...")
    coord_encontrada_ocr = selecionar_guia_planilha_via_toggle_e_nome(
        nome_guia=nome_guia,
        coord_toggle_todas_guias=t_tuple,
        coord_item_no_toggle=None,
        coord_guia_fallback=None,
    )
    if coord_encontrada_ocr is not None:
        padrao_pagina["coord_item_no_toggle"] = [int(coord_encontrada_ocr[0]), int(coord_encontrada_ocr[1])]
        i_tuple = coord_encontrada_ocr
    else:
        print(
            f"   👉 Agora, com a lista do Toggle aberta (ou na barra inferior), posicione o mouse sobre "
            f"'{nome_guia}' e aperte [F8]:"
        )
        mover_e_clicar(x=int(tx), y=int(ty), cliques=1, botao="left")
        time.sleep(0.45)
        ix, iy = aguardar_tecla_f8()
        i_tuple = (int(ix), int(iy))
        padrao_pagina["coord_item_no_toggle"] = [int(ix), int(iy)]
        mover_e_clicar(x=int(ix), y=int(iy), cliques=1, botao="left")
        time.sleep(0.55)

    salvar_mapeamento(mapeamento)
    print("=" * 72)
    return t_tuple, i_tuple, g_tuple


def _garantir_coord_guia_planilha(
    padrao_pagina: Dict[str, Any],
    mapeamento: Dict[str, Any],
    nome_guia_padrao: str,
    solicitar_se_ausente: bool = True,
    forcar_recalibracao: bool = False,
) -> Optional[Tuple[int, int]]:
    """
    Mantido por compatibilidade: retorna coordenada de fallback ou aciona a calibração do Toggle + guia.
    """
    t_tuple, i_tuple, g_tuple = _garantir_coord_toggle_ou_guia_planilha(
        padrao_pagina=padrao_pagina,
        mapeamento=mapeamento,
        nome_guia_padrao=nome_guia_padrao,
        solicitar_se_ausente=solicitar_se_ausente,
        forcar_recalibracao=forcar_recalibracao,
    )
    return g_tuple or i_tuple or t_tuple


def garantir_guia_correta_na_planilha_e_copiar(
    padrao_bloco_ou_pagina: Dict[str, Any],
    mapeamento: Dict[str, Any],
    coord_celula_planilha: Tuple[int, int],
    nomes_esperados: List[str],
    nome_guia: str,
    rotulo_etapa: str,
    linhas_atuais: Optional[List[Dict[str, Any]]] = None,
    verificar_aba_aberta_na_planilha: bool = True,
) -> List[Dict[str, Any]]:
    """
    Implementa a regra oficial do 1º Passo em TODAS as fases e blocos:
    1. PRIMEIRO verifica na planilha se JÁ ESTÁ na guia certa (`nome_guia`):
       - Se `linhas_atuais` já foram lidas e correspondem a `nome_guia`, confirma imediatamente.
       - Caso contrário, se `verificar_aba_aberta_na_planilha=True`, faz uma leitura rápida da aba aberta
         na planilha e verifica se as colunas batem com `nomes_esperados` (`_confirmar_correspondencia_guia_planilha`).
    2. Se JÁ ESTIVER na guia certa: não clica em outra aba e segue direto com os dados copiados!
    3. Se NÃO ESTIVER na guia certa: usa o Toggle de Todas as Guias (`≡`) + Nome de Referência (`nome_guia`)
       para selecionar a guia correta na planilha, e SÓ DEPOIS faz a cópia (`ESC -> Ctrl+A -> Ctrl+C`)
       e confirma os dados.
    """
    print("\n" + "=" * 72)
    print(f"🔎 [1º PASSO — VERIFICAR GUIA DA PLANILHA] Etapa: {rotulo_etapa} | Guia esperada: '{nome_guia}'")

    if linhas_atuais and _confirmar_correspondencia_guia_planilha(linhas_atuais, nomes_esperados, nome_guia):
        print(
            f"   ✅ A planilha JÁ está na guia certa ('{nome_guia}') e os dados ({len(linhas_atuais)} linha(s)) "
            "já estão prontos!"
        )
        print("=" * 72)
        return linhas_atuais

    if verificar_aba_aberta_na_planilha:
        print(f"   📋 Verificando na planilha se a guia aberta atualmente já é '{nome_guia}'...")
        linhas_teste = _copiar_tabela_da_guia_atual_na_planilha(
            coord_celula_planilha=coord_celula_planilha,
            nomes_esperados=nomes_esperados,
            nome_guia=nome_guia,
        )
        if _confirmar_correspondencia_guia_planilha(linhas_teste, nomes_esperados, nome_guia):
            print(
                f"   ✅ Confirmado! A planilha JÁ estava na guia certa ('{nome_guia}')! "
                "Seguindo direto para a verificação da guia no sistema..."
            )
            print("=" * 72)
            return linhas_teste
        print(
            f"   ℹ️  A planilha NÃO está na guia '{nome_guia}'. "
            f"Selecionando '{nome_guia}' via Toggle de Todas as Guias ('≡') + Nome de Referência..."
        )

    t_tuple, i_tuple, g_tuple = _garantir_coord_toggle_ou_guia_planilha(
        padrao_pagina=padrao_bloco_ou_pagina,
        mapeamento=mapeamento,
        nome_guia_padrao=nome_guia,
        solicitar_se_ausente=True,
        forcar_recalibracao=False,
    )

    linhas_copiadas = trocar_guia_planilha_e_copiar_tabela(
        coord_guia_planilha=g_tuple,
        coord_celula_planilha=coord_celula_planilha,
        nomes_esperados=nomes_esperados,
        nome_guia=nome_guia,
        rotulo_ligacao=rotulo_etapa,
        coord_toggle_todas_guias=t_tuple,
        coord_item_no_toggle=i_tuple,
    )

    if not _confirmar_correspondencia_guia_planilha(linhas_copiadas, nomes_esperados, nome_guia):
        print(
            f"⚠️  Os dados copiados ainda não correspondem à guia '{nome_guia}'. "
            "Vamos calibrar o Toggle de Todas as Guias ('≡') e selecionar a guia correta:"
        )
        t_tuple, i_tuple, g_tuple = _garantir_coord_toggle_ou_guia_planilha(
            padrao_pagina=padrao_bloco_ou_pagina,
            mapeamento=mapeamento,
            nome_guia_padrao=nome_guia,
            solicitar_se_ausente=True,
            forcar_recalibracao=True,
        )
        linhas_copiadas = _copiar_tabela_da_guia_atual_na_planilha(
            coord_celula_planilha=coord_celula_planilha,
            nomes_esperados=nomes_esperados,
            nome_guia=nome_guia,
        )

    return linhas_copiadas


def _selecionar_guia_sistema_apos_copiar_tabela(
    padrao_pagina: Dict[str, Any],
    mapeamento: Dict[str, Any],
    nome_pagina_ativa: str,
    step_index_padrao: Optional[int] = None,
    titulo_coluna_confirmacao: Optional[str] = None,
) -> Tuple[int, int]:
    """
    Executado IMEDIATAMENTE APÓS confirmar/copiar a tabela na planilha:
    1. PRIMEIRO verifica no sistema se ele já está na guia correspondente (`titulo_coluna_confirmacao`).
    2. Se não estiver, clica na guia correta dentro de `<div class="nav flex-column wizard-nav">`
       (ex: `3. Diagnóstico`, `4. Metas e Planos`, `5. Capacitação e Comunicação`, `6. Monitoramento e Revisão`)
       e confirma que a coluna/seção esperada está visível antes de colar.
    """
    nome_guia_sis = str(padrao_pagina.get("nome_guia_sistema") or nome_pagina_ativa)
    norm_sis = remover_acentos(nome_guia_sis)
    if "METAS" in norm_sis and "OBRIGATORIO" in norm_sis:
        nome_guia_sis = "4. Metas e Planos"
    elif "CAPACITACAO" in norm_sis:
        nome_guia_sis = "5. Capacitação e Comunicação"
    elif "MONITORAMENTO" in norm_sis:
        nome_guia_sis = "6. Monitoramento e Revisão"

    step_idx = padrao_pagina.get("step_index_sistema")
    if step_idx is None:
        step_idx = step_index_padrao

    if titulo_coluna_confirmacao is None:
        if "METAS" in norm_sis:
            titulo_coluna_confirmacao = "Metas e ações"
        elif "DIAGNOSTICO" in norm_sis:
            titulo_coluna_confirmacao = "Forças"
        elif "CAPACITACAO" in norm_sis:
            titulo_coluna_confirmacao = "Capacitação"
        elif "MONITORAMENTO" in norm_sis:
            titulo_coluna_confirmacao = "Monitoramento"

    coord_salva_raw = padrao_pagina.get("coord_guia_sistema")
    coord_salva = (
        (int(coord_salva_raw[0]), int(coord_salva_raw[1]))
        if isinstance(coord_salva_raw, (list, tuple)) and len(coord_salva_raw) >= 2
        else None
    )

    coord_clicada = selecionar_guia_wizard_no_sistema(
        nome_guia_sistema=nome_guia_sis,
        coord_guia_sistema=coord_salva,
        step_index=int(step_idx) if step_idx is not None else None,
        titulo_coluna_confirmacao=titulo_coluna_confirmacao,
    )

    if coord_salva is None and coord_clicada:
        padrao_pagina["coord_guia_sistema"] = [int(coord_clicada[0]), int(coord_clicada[1])]
        padrao_pagina["nome_guia_sistema"] = nome_guia_sis
        if step_idx is not None:
            padrao_pagina["step_index_sistema"] = int(step_idx)
        salvar_mapeamento(mapeamento)

    return coord_clicada


def _avancar_para_proxima_etapa_do_sistema(
    mapeamento: Dict[str, Any],
    padrao_pagina_atual: Dict[str, Any],
) -> Optional[Tuple[int, int]]:
    """
    Ao terminar todas as linhas e blocos da página atual no sistema, clica no botão
    `<button type="button" class="btn btn-primary" data-step-action="next">Próxima etapa</button>`
    para avançar automaticamente para a próxima página da sequência.
    """
    coord_btn_raw = (
        padrao_pagina_atual.get("coord_botao_proxima_etapa")
        or mapeamento.get("coord_botao_proxima_etapa")
    )
    coord_btn = (
        (int(coord_btn_raw[0]), int(coord_btn_raw[1]))
        if isinstance(coord_btn_raw, (list, tuple)) and len(coord_btn_raw) >= 2 and coord_btn_raw[0] is not None
        else None
    )
    coord_clicada = clicar_botao_proxima_etapa_no_sistema(coord_botao_salva=coord_btn)
    if coord_btn is None and coord_clicada is not None:
        mapeamento["coord_botao_proxima_etapa"] = [int(coord_clicada[0]), int(coord_clicada[1])]
        padrao_pagina_atual["coord_botao_proxima_etapa"] = [int(coord_clicada[0]), int(coord_clicada[1])]
        salvar_mapeamento(mapeamento)
    return coord_clicada


def _executar_roteiro_de_bloco(
    linhas: List[Dict[str, Any]],
    nome_bloco: str,
    padrao_bloco: Dict[str, Any],
    mapeamento: Dict[str, Any],
    params: Dict[str, Any],
    mapa_html: Dict[str, Dict[str, Any]],
    ir_para_topo_no_inicio: bool = True,
    rolar_ate_o_final_no_inicio: bool = False,
) -> bool:
    """
    Executa todos os passos (`coluna_independente`, `grupo_codependente`, `campo_unico`) de um bloco da página.
    - No Bloco 1 (SWOT): começa no topo para 'Forças' e 'Fraquezas', e logo após 'Forças' e 'Fraquezas'
      rola a tela até o final (`rolar_pagina_ate_o_final`) para posicionar 'Oportunidades' e 'Ameaças'.
    - No Bloco 2 ('Inventário de necessidades'): rola a página até o final novamente (`rolar_ate_o_final_no_inicio=True`)
      para posicionar 'Inventário de necessidades' e preencher os 4 campos de cada linha.
    """
    campos_detalhados = (
        padrao_bloco.get("campos")
        or mapeamento.get("campos_detalhados")
        or {}
    )
    caixas_bloqueadas_pagina = padrao_bloco.get("caixas_bloqueadas") or []

    passo_y_sistema = int(params["passo_y_nova_caixa"])
    xn, yn = params["ponto_neutro"]
    xn_seguro = max(620, min(1800, int(xn)))
    yn_seguro = max(195, min(350, int(yn)))

    colunas_com_itens = extrair_colunas_da_tabela(linhas, campos_detalhados)
    mapa_itens_por_coluna: Dict[str, List[Tuple[int, str]]] = {
        remover_acentos(nome): itens for nome, itens in colunas_com_itens
    }

    roteiro_passos = _montar_roteiro_passos_da_pagina(
        padrao_pagina=padrao_bloco,
        colunas_com_itens=colunas_com_itens,
        campos_detalhados=campos_detalhados,
    )

    descricoes_roteiro = [
        f"{p.get('nome_passo')} ({'GRUPO ' + str(len(p.get('campos') or [])) + ' CAMPOS' if p.get('tipo_passo') == 'grupo_codependente' else 'COLUNA'})"
        for p in roteiro_passos
    ]

    print("\n" + "=" * 72)
    print(f"📄 BLOCO ATIVO: '{nome_bloco}'")
    print(f"🚀 ROTEIRO PASSO A PASSO DO BLOCO ({len(roteiro_passos)} etapa(s))")
    print(f"📌 Sequência: {' -> '.join(descricoes_roteiro)}")
    print(
        f"⏱️  Espera após clicar em 'Adicionar': {ESPERA_APOS_ADICIONAR:.2f}s "
        f"| Caixas mapeadas no bloco: {len(campos_detalhados)}"
    )
    print("💡 Pressione [ESC] a qualquer momento para abortar com segurança.")
    print("=" * 72)

    if ir_para_topo_no_inicio:
        ir_para_topo_da_pagina(xn_seguro, yn_seguro)
        ja_rolou_para_baixo = False
    elif rolar_ate_o_final_no_inicio:
        print(f"  📜 Posicionando '{nome_bloco}': rolando a página até o final...")
        rolar_pagina_ate_o_final(xn_seguro, yn_seguro)
        ja_rolou_para_baixo = True
    else:
        ja_rolou_para_baixo = True

    qtd_itens_preenchidos_por_coluna: Dict[str, int] = {}

    for idx_passo, passo in enumerate(roteiro_passos):
        if verificar_aborto():
            return False

        tipo_passo = str(passo.get("tipo_passo") or "coluna_independente")

        # Caso 1: Grupo Codependente (ex: Bloco 2 'Inventário de necessidades' — 4 campos antes de 'Adicionar necessidade')
        if tipo_passo == "grupo_codependente":
            ok_grupo, ja_rolou_para_baixo = _executar_passo_grupo_codependente(
                passo=passo,
                linhas=linhas,
                campos_detalhados=campos_detalhados,
                caixas_bloqueadas_pagina=caixas_bloqueadas_pagina,
                mapa_html=mapa_html,
                xn=xn_seguro,
                ja_rolou_para_baixo=ja_rolou_para_baixo,
            )
            if not ok_grupo:
                return False
            continue

        # Caso 2: Coluna Independente ou Campo Único
        campos_passo = passo.get("campos") or []
        nome_campo = str(campos_passo[0]) if campos_passo else str(passo.get("nome_passo") or "")
        if not nome_campo:
            continue

        info_calibrada = _buscar_info_campo_calibrado(nome_campo, campos_detalhados)
        tipo_campo = str(info_calibrada.get("tipo", "caixa_texto")).lower()
        if tipo_campo == "bloqueada":
            print(
                f"\n🛡️  Caixa '{nome_campo}' está marcada como PROTEGIDA/BLOQUEADA neste bloco (não será alterada)."
            )
            continue

        itens_coluna = mapa_itens_por_coluna.get(remover_acentos(nome_campo), [])
        if not itens_coluna:
            for idx_l, linha in enumerate(linhas, start=1):
                val_l = buscar_valor_da_coluna(linha, nome_campo)
                if val_l:
                    itens_coluna.append((idx_l, val_l))

        if not itens_coluna:
            print(f"\n⚪ Passo {idx_passo + 1}/{len(roteiro_passos)} ('{nome_campo}'): nenhum valor para preencher.")
            continue

        if LIMITE_LINHAS_TESTE is not None and LIMITE_LINHAS_TESTE > 0 and len(itens_coluna) > LIMITE_LINHAS_TESTE:
            total_orig_col = len(itens_coluna)
            itens_coluna = itens_coluna[:LIMITE_LINHAS_TESTE]
            print(
                f"  🔢 [LIMITE DE TESTE ATIVO] Limitando coluna '{nome_campo}' para as primeiras "
                f"{len(itens_coluna)} de {total_orig_col} linha(s)!"
            )

        campos_simples = mapeamento.get("campos", {})
        if "sistema_campo_linha1" in info_calibrada and info_calibrada["sistema_campo_linha1"]:
            coord_campo_salva = tuple(info_calibrada["sistema_campo_linha1"])
        elif nome_campo in campos_simples and campos_simples[nome_campo]:
            coord_campo_salva = tuple(campos_simples[nome_campo])
        else:
            coord_campo_salva = None
        coord_btn_salva = (
            tuple(info_calibrada["botao_adicionar"])
            if info_calibrada.get("botao_adicionar")
            else None
        )

        chave_norm = remover_acentos(nome_campo)
        info_h = mapa_html.get(chave_norm)
        if info_h and info_h.get("tag") == "select" and tipo_campo == "caixa_texto":
            tipo_campo = "barra_suspensa"

        eh_quadrante_inferior = (
            chave_norm in ("OPORTUNIDADES", "OPORTUNIDADE", "AMEACAS", "AMEACA")
            or (coord_btn_salva is not None and int(coord_btn_salva[1]) > 600)
            or (info_h is not None and int(info_h.get("ordem", 0)) >= 2 and idx_passo >= 2)
        )

        if coord_btn_salva is None and params.get("botao_ref"):
            bx_ref, by_ref = int(params["botao_ref"][0]), int(params["botao_ref"][1])
            if chave_norm in ("FRAQUEZAS", "FRAQUEZA", "AMEACAS", "AMEACA"):
                coord_btn_salva = (bx_ref + 590 if bx_ref < 1400 else bx_ref, by_ref)
            elif chave_norm in ("FORCAS", "FORCA", "OPORTUNIDADES", "OPORTUNIDADE"):
                coord_btn_salva = (bx_ref if bx_ref < 1400 else bx_ref - 590, by_ref)

        zonas_proibidas_coluna = _obter_zonas_proibidas_da_pagina(
            nomes_campos_atuais=[nome_campo],
            campos_detalhados=campos_detalhados,
            caixas_bloqueadas=caixas_bloqueadas_pagina,
            coord_atual_ref=coord_campo_salva,
        )

        etiqueta_tipo = "BARRA SUSPENSA / TOGGLE" if tipo_campo == "barra_suspensa" else "CAIXA DE TEXTO"
        print(
            f"\n📌 PASSO {idx_passo + 1}/{len(roteiro_passos)} (COLUNA INDEPENDENTE): "
            f"'{nome_campo}' [{etiqueta_tipo}] — preenchendo todos os {len(itens_coluna)} valor(es) do início ao fim"
        )

        coord_caixa_atual, coord_btn_atual, ja_rolou_para_baixo = localizar_card_com_ctrl_f_e_visao(
            nome_campo=nome_campo,
            coord_campo_salva=coord_campo_salva,
            coord_btn_salva=coord_btn_salva,
            eh_quadrante_inferior=eh_quadrante_inferior,
            ja_rolou_para_baixo=ja_rolou_para_baixo,
        )

        if tipo_passo == "campo_unico" or info_calibrada.get("tem_adicionar") is False:
            coord_btn_atual = None

        y_lim_bloco_viz_raw = (
            passo.get("y_limite_bloco_vizinho")
            or info_calibrada.get("y_limite_bloco_vizinho")
            or padrao_bloco.get("y_limite_bloco_vizinho")
        )
        y_lim_bloco_viz = int(y_lim_bloco_viz_raw) if y_lim_bloco_viz_raw is not None else None

        ticks_rolados_na_coluna = 0

        for idx_item, (num_linha_planilha, valor) in enumerate(itens_coluna):
            if verificar_aborto():
                return False

            try:
                if tipo_campo != "barra_suspensa":
                    coord_ajustada_col, esta_vazia_col, bloq_col = ajustar_coordenada_para_caixa_vazia(
                        coord_alvo=coord_caixa_atual,
                        coord_btn_add=coord_btn_atual,
                        zonas_proibidas=zonas_proibidas_coluna,
                        y_limite_bloco_vizinho=y_lim_bloco_viz,
                        permitir_avancar_abaixo=True,
                        nome_campo=nome_campo,
                    )
                    if not bloq_col:
                        if abs(int(coord_ajustada_col[1]) - int(coord_caixa_atual[1])) > 25 and esta_vazia_col:
                            print(
                                f"  🔍 [SÓ CAIXA VAZIA] Caixa anterior em {coord_caixa_atual} já ocupada; "
                                f"avançando automaticamente para a caixa vazia em {coord_ajustada_col}!"
                            )
                        coord_caixa_atual = coord_ajustada_col

                ok_item = _preencher_ou_selecionar_valor_unico(
                    nome_campo=nome_campo,
                    valor=valor,
                    coord_alvo=coord_caixa_atual,
                    info_calibrada=info_calibrada,
                    mapa_html=mapa_html,
                    num_linha_planilha=num_linha_planilha,
                    idx_caixa=idx_item + 1,
                    total_caixas=len(itens_coluna),
                    coord_btn_add=coord_btn_atual,
                    zonas_proibidas=zonas_proibidas_coluna,
                    y_limite_bloco_vizinho=y_lim_bloco_viz,
                    permitir_avancar_abaixo=False,
                )
                if not ok_item:
                    return False

                tem_proximo_item = (idx_item + 1) < len(itens_coluna)
                if tem_proximo_item:
                    if coord_btn_atual is not None:
                        num_clique_add = idx_item + 1
                        eh_repeater_etapa_5 = (
                            chave_norm
                            in (
                                "TREINAMENTOS TECNICOS",
                                "ACOES DE RETENCAO DE TALENTOS",
                                "ACOES MACRO",
                                "ACOES MICRO",
                                "ACOES EXTRAS",
                            )
                            or bool(passo.get("scroll_down_apos_adicionar", False))
                        )
                        forcar_scroll_4 = (num_clique_add % 4 == 0) or (
                            eh_repeater_etapa_5 and (coord_caixa_atual[1] > 760 or num_clique_add >= 2)
                        )
                        msg_scroll_4 = (
                            " + scroll down dinâmico para focar o novo campo recém-criado"
                            if eh_repeater_etapa_5 and forcar_scroll_4
                            else (
                                " + scroll down padronizado (a cada 4 cliques em Adicionar)"
                                if forcar_scroll_4
                                else ""
                            )
                        )
                        print(
                            f"  ➕ [{nome_campo}] Clicando em 'Adicionar' (#{num_clique_add}) em {coord_btn_atual}, "
                            f"aguardando {ESPERA_APOS_ADICIONAR:.2f}s{msg_scroll_4} e focando o novo campo vazio #{idx_item + 2} (índice [{idx_item + 1}] / :last-of-type)..."
                        )
                        coord_caixa_atual, coord_btn_atual, ticks_passo = (
                            clicar_adicionar_e_avancar_caixa(
                                coord_caixa_atual=coord_caixa_atual,
                                coord_btn_atual=coord_btn_atual,
                                passo_y_padrao=passo_y_sistema,
                                zonas_proibidas=zonas_proibidas_coluna,
                                forcar_scroll_padrao=forcar_scroll_4,
                            )
                        )
                        ticks_rolados_na_coluna += ticks_passo
                    else:
                        coord_caixa_atual = (
                            coord_caixa_atual[0],
                            min(990, coord_caixa_atual[1] + passo_y_sistema),
                        )
            except Exception as exc:
                print(f"  ❌ Erro ao preencher '{nome_campo}' (Linha {num_linha_planilha}): {exc}")

            time.sleep(ESPERA_CURTA)

        qtd_itens_preenchidos_por_coluna[chave_norm] = len(itens_coluna)

        print(
            f"  ✅ Coluna '{nome_campo}' 100% concluída "
            f"({len(itens_coluna)}/{len(itens_coluna)} valores preenchidos)!"
        )

        # Transição inteligente de rolagem para a PRÓXIMA coluna da página:
        if (idx_passo + 1) < len(roteiro_passos):
            prox_passo = roteiro_passos[idx_passo + 1]
            prox_campos = prox_passo.get("campos") or []
            prox_nome = str(prox_campos[0]) if prox_campos else str(prox_passo.get("nome_passo") or "")
            prox_norm = remover_acentos(prox_nome)

            if prox_norm in (
                "FRAQUEZAS",
                "FRAQUEZA",
                "ACOES DE RETENCAO DE TALENTOS",
                "RETENCAO DE TALENTOS",
            ):
                # 'Fraquezas' (Pág 3) e 'Ações de retenção de talentos' (Pág 5) ficam no topo direito: volta para o topo!
                print(f"  📜 [Transição -> '{prox_nome}'] Voltando ao topo da página para iniciar a coluna da direita...")
                ir_para_topo_da_pagina(xn_seguro, yn_seguro)
                ja_rolou_para_baixo = False
            elif prox_norm in ("OPORTUNIDADES", "OPORTUNIDADE", "ACOES MACRO"):
                # Rola a tela até o final para posicionar os blocos inferiores ('Oportunidades'/'Ameaças' ou 'Ações macro'/'micro'/'extras')
                print(
                    f"  📜 [Transição -> '{prox_nome}'] Rolando a tela até o final para posicionar '{prox_nome}'..."
                )
                rolar_pagina_ate_o_final(xn_seguro, yn_seguro)
                ja_rolou_para_baixo = True
            elif prox_norm in ("AMEACAS", "AMEACA", "ACOES MICRO", "ACOES EXTRAS"):
                # Colunas lado a lado na parte inferior: devolve os ticks rolados na coluna anterior para alinhar com o cabeçalho
                if ticks_rolados_na_coluna > 0:
                    print(
                        f"  📜 [Transição -> '{prox_nome}'] Subindo {ticks_rolados_na_coluna} tick(s) "
                        f"para alinhar com o cabeçalho de '{prox_nome}'..."
                    )
                    rolar_pagina(ticks_rolados_na_coluna, x=1230, y=500)
                    time.sleep(0.40)
            else:
                if ticks_rolados_na_coluna > 0:
                    rolar_pagina(ticks_rolados_na_coluna, x=1230, y=500)
                    time.sleep(0.35)

    return True


def _executar_blocos_adicionais_da_pagina(
    nome_pagina: str,
    padrao_pagina: Dict[str, Any],
    mapeamento: Dict[str, Any],
    params: Dict[str, Any],
    mapa_html: Dict[str, Dict[str, Any]],
    step_index_sistema: int,
) -> bool:
    """
    Executa em sequência quaisquer blocos adicionais configurados na página (`bloco_2`, `blocos_adicionais`),
    sempre seguindo a regra:
    1º Verifica na planilha se já está na guia certa; se não estiver, seleciona pelo Toggle '≡' + nome e copia.
    2º Verifica no sistema se está na guia certa.
    3º Preenche todas as linhas do bloco e passa para o bloco seguinte.
    """
    lista_blocos: List[Tuple[int, Dict[str, Any]]] = []
    b2 = padrao_pagina.get("bloco_2")
    if isinstance(b2, dict) and b2.get("ativo", True) and (b2.get("campos") or b2.get("passos_pagina")):
        lista_blocos.append((2, b2))

    blocos_extras = padrao_pagina.get("blocos_adicionais")
    if isinstance(blocos_extras, list):
        for idx_extra, b_extra in enumerate(blocos_extras, start=3):
            if isinstance(b_extra, dict) and b_extra.get("ativo", True) and (b_extra.get("campos") or b_extra.get("passos_pagina")):
                lista_blocos.append((idx_extra, b_extra))

    for num_b, info_b in lista_blocos:
        nome_b = str(info_b.get("nome_bloco") or f"Bloco {num_b}")
        nome_guia_b = str(info_b.get("nome_guia_planilha") or nome_b)
        campos_b = info_b.get("campos") or {}
        nomes_esp_b = list(campos_b.keys())
        if not nomes_esp_b:
            continue

        coord_cel_b = info_b.get("coord_celula_planilha") or params["planilha"]
        linhas_b = garantir_guia_correta_na_planilha_e_copiar(
            padrao_bloco_ou_pagina=info_b,
            mapeamento=mapeamento,
            coord_celula_planilha=(int(coord_cel_b[0]), int(coord_cel_b[1])),
            nomes_esperados=nomes_esp_b,
            nome_guia=nome_guia_b,
            rotulo_etapa=f"{nome_pagina} — BLOCO {num_b} ('{nome_guia_b}')",
            linhas_atuais=None,
            verificar_aba_aberta_na_planilha=True,
        )
        if not linhas_b:
            continue

        _selecionar_guia_sistema_apos_copiar_tabela(
            padrao_pagina=padrao_pagina,
            mapeamento=mapeamento,
            nome_pagina_ativa=nome_pagina,
            step_index_padrao=step_index_sistema,
            titulo_coluna_confirmacao=str(info_b.get("titulo_coluna_confirmacao") or nome_b),
        )

        ok_b = _executar_roteiro_de_bloco(
            linhas=linhas_b,
            nome_bloco=f"{nome_pagina} — Bloco {num_b} ({nome_b})",
            padrao_bloco=info_b,
            mapeamento=mapeamento,
            params=params,
            mapa_html=mapa_html,
            ir_para_topo_no_inicio=bool(info_b.get("ir_para_topo_no_inicio", False)),
            rolar_ate_o_final_no_inicio=bool(info_b.get("rolar_ate_o_final_no_inicio", True)),
        )
        if not ok_b:
            return False

    return True


def _garantir_estrutura_blocos_fases_4_5_6(
    mapeamento: Dict[str, Any],
    params: Dict[str, Any],
) -> None:
    """
    Garante que os blocos adicionais da Página 4 (Bloco 2: '4.2. Plano Orçamentário' e Bloco 3: '4.3. Gestão de Riscos'),
    da Página 5 (Passo A: '5.1. Capacitação' e Passo B: '5.3. Ações Específicas') e da Página 6
    ('6. Monitoramento, Anexos e Revisão' — 'Indicadores de sucesso (KPIs)' via memória da Etapa 4)
    estejam estruturados em `padroes_paginas`, mantendo 100% intactos os blocos já calibrados
    da Página 3 ('3.1. SWOT' e '3.2. Inventário de necessidades') e o Bloco 1 da Página 4 ('4.1. Metas e Ações').
    """
    padroes = mapeamento.setdefault("padroes_paginas", {})
    coord_plan_padrao = list(params.get("planilha") or [-1534, 588])
    coord_toggle_raiz = mapeamento.get("coord_toggle_todas_guias")

    # =========================================================================
    # PÁGINA 4: Mantém Bloco 1 ('4.1. Metas e Ações') intacto e garante Bloco 2 e Bloco 3
    # =========================================================================
    p4 = padroes.setdefault(
        "4. Metas e Planos Obrigatórios",
        {"nome_pagina": "4. Metas e Planos Obrigatórios", "campos": {}},
    )
    coord_cel_p4 = list(p4.get("coord_celula_planilha") or coord_plan_padrao)

    # Bloco 2 da Página 4: Guia '4.2. Plano Orçamentário' -> 'Plano orçamentário e de investimentos'
    if not isinstance(p4.get("bloco_2"), dict) or not (p4["bloco_2"].get("campos")):
        campos_b2_p4: Dict[str, Any] = {}
        definicao_b2_p4 = [
            ("AÇÃO (ID)", "AÇÃO (ID)", "caixa_texto", None, [755, 620]),
            ("DESCRIÇÃO DO CUSTO", "DESCRIÇÃO DO CUSTO", "caixa_texto", None, [1065, 620]),
            ("FONTE DE RECURSO", "FONTE DE RECURSO", "caixa_texto", None, [1500, 620]),
            ("ANO 1", "ANO 1", "caixa_texto", None, [805, 709]),
            ("ANO 2", "ANO 2", "caixa_texto", None, [1065, 709]),
            ("ANO 3", "ANO 3", "caixa_texto", None, [1325, 709]),
            ("ANO 4", "ANO 4", "caixa_texto", None, [1585, 709]),
        ]
        for nome_c, col_plan, tipo_c, opcoes_c, coord_sug in definicao_b2_p4:
            campos_b2_p4[nome_c] = {
                "coluna_planilha": col_plan,
                "planilha_linha1": list(coord_cel_p4),
                "sistema_campo_linha1": list(coord_sug),
                "botao_adicionar": [1769, 510],
                "tipo": tipo_c,
                "tem_adicionar": True,
                "passo_y_grupo": 188,
            }
        p4["bloco_2"] = {
            "nome_bloco": "Plano orçamentário e de investimentos",
            "nome_guia_planilha": "4.2. Plano Orçamentário",
            "titulo_coluna_confirmacao": "Plano orçamentário",
            "ativo": True,
            "coord_toggle_todas_guias": coord_toggle_raiz,
            "coord_celula_planilha": list(coord_cel_p4),
            "ir_para_topo_no_inicio": False,
            "rolar_ate_o_final_no_inicio": False,
            "ancorado_no_topo": True,
            "campos": campos_b2_p4,
            "caixas_bloqueadas": [],
            "passos_pagina": [
                {
                    "tipo_passo": "grupo_codependente",
                    "nome_passo": "Plano orçamentário e de investimentos",
                    "titulo_card_busca": "Plano orçamentário e de investimentos",
                    "texto_botao_adicionar": "Adicionar linha",
                    "data_repeater_add": "budget-plan",
                    "ancorado_no_topo": True,
                    "campos": list(COLUNAS_PLANO_ORCAMENTARIO),
                    "botao_adicionar": [1769, 510],
                    "passo_y_grupo": 188,
                }
            ],
        }

    # Bloco 3 da Página 4: Guia '4.3. Gestão de Riscos' -> 'Plano de gestão de riscos'
    # Nota HTML: Todos os 6 campos de risk-plan são <input type="text" class="form-control">!
    blocos_extras_p4 = p4.get("blocos_adicionais")
    if not isinstance(blocos_extras_p4, list) or not blocos_extras_p4:
        campos_b3_p4: Dict[str, Any] = {}
        definicao_b3_p4 = [
            ("RISCO IDENTIFICADO", "RISCO IDENTIFICADO", "caixa_texto", None, [815, 620]),
            ("PROBABILIDADE", "PROBABILIDADE", "caixa_texto", None, [1055, 620]),
            ("IMPACTO", "IMPACTO", "caixa_texto", None, [1245, 620]),
            ("NÍVEL", "NÍVEL", "caixa_texto", None, [1530, 620]),
            ("AÇÃO DE MITIGAÇÃO", "AÇÃO DE MITIGAÇÃO", "caixa_texto", None, [1020, 709]),
            ("RESPONSÁVEL PELO RISCO", "RESPONSÁVEL PELO RISCO", "caixa_texto", None, [1610, 709]),
        ]
        for nome_c, col_plan, tipo_c, opcoes_c, coord_sug in definicao_b3_p4:
            reg_b3: Dict[str, Any] = {
                "coluna_planilha": col_plan,
                "planilha_linha1": list(coord_cel_p4),
                "sistema_campo_linha1": list(coord_sug),
                "botao_adicionar": [1769, 510],
                "tipo": tipo_c,
                "tem_adicionar": True,
                "passo_y_grupo": 188,
            }
            campos_b3_p4[nome_c] = reg_b3

        p4["blocos_adicionais"] = [
            {
                "nome_bloco": "Plano de gestão de riscos",
                "nome_guia_planilha": "4.3. Gestão de Riscos",
                "titulo_coluna_confirmacao": "Plano de gestão de riscos",
                "ativo": True,
                "coord_toggle_todas_guias": coord_toggle_raiz,
                "coord_celula_planilha": list(coord_cel_p4),
                "ir_para_topo_no_inicio": False,
                "rolar_ate_o_final_no_inicio": False,
                "ancorado_no_topo": True,
                "campos": campos_b3_p4,
                "caixas_bloqueadas": [],
                "passos_pagina": [
                    {
                        "tipo_passo": "grupo_codependente",
                        "nome_passo": "Plano de gestão de riscos",
                        "titulo_card_busca": "Plano de gestão de riscos",
                        "texto_botao_adicionar": "Adicionar risco",
                        "data_repeater_add": "risk-plan",
                        "ancorado_no_topo": True,
                        "campos": list(COLUNAS_GESTAO_RISCOS),
                        "botao_adicionar": [1769, 510],
                        "passo_y_grupo": 188,
                    }
                ],
            }
        ]

    # =========================================================================
    # PÁGINA 5: '5. Capacitação, Comunicação e Ações Específicas'
    # Passo A (Guia '5.1. Capacitação') + Passo B (Guia '5.3. Ações Específicas')
    # =========================================================================
    p5 = padroes.setdefault(
        "5. Capacitação, Comunicação e Ações Específicas",
        {"nome_pagina": "5. Capacitação, Comunicação e Ações Específicas", "campos": {}},
    )
    if not p5.get("campos"):
        p5["nome_bloco_1"] = "5. Capacitação — Passo A (Guia 5.1. Capacitação)"
        p5["nome_guia_planilha"] = "5.1. Capacitação"
        p5["nome_guia_sistema"] = "5. Capacitação e Comunicação"
        p5["step_index_sistema"] = 4
        p5["titulo_coluna_confirmacao"] = "Treinamentos técnicos"
        p5["coord_toggle_todas_guias"] = coord_toggle_raiz
        p5["coord_celula_planilha"] = list(coord_cel_p4)
        p5["campos"] = {
            "TREINAMENTOS TÉCNICOS": {
                "coluna_planilha": "TREINAMENTOS TÉCNICOS",
                "planilha_linha1": list(coord_cel_p4),
                "sistema_campo_linha1": [735, 685],
                "botao_adicionar": [1174, 622],
                "tipo": "caixa_texto",
                "tem_adicionar": True,
                "data_repeater_add": "technical-trainings",
            },
            "AÇÕES DE RETENÇÃO DE TALENTOS": {
                "coluna_planilha": "AÇÕES DE RETENÇÃO DE TALENTOS",
                "planilha_linha1": list(coord_cel_p4),
                "sistema_campo_linha1": [1386, 685],
                "botao_adicionar": [1788, 622],
                "tipo": "caixa_texto",
                "tem_adicionar": True,
                "data_repeater_add": "talent-actions",
            },
        }
        p5["passos_pagina"] = [
            {
                "tipo_passo": "coluna_independente",
                "nome_passo": "TREINAMENTOS TÉCNICOS",
                "campos": ["TREINAMENTOS TÉCNICOS"],
                "data_repeater_add": "technical-trainings",
                "scroll_down_apos_adicionar": True,
            },
            {
                "tipo_passo": "coluna_independente",
                "nome_passo": "AÇÕES DE RETENÇÃO DE TALENTOS",
                "campos": ["AÇÕES DE RETENÇÃO DE TALENTOS"],
                "data_repeater_add": "talent-actions",
                "scroll_down_apos_adicionar": True,
            },
        ]

    if not isinstance(p5.get("bloco_2"), dict) or not (p5["bloco_2"].get("campos")):
        p5["bloco_2"] = {
            "nome_bloco": "Passo B — Ações Específicas (5.3. Ações Específicas)",
            "nome_guia_planilha": "5.3. Ações Específicas",
            "titulo_coluna_confirmacao": "Ações macro",
            "ativo": True,
            "coord_toggle_todas_guias": coord_toggle_raiz,
            "coord_celula_planilha": list(coord_cel_p4),
            "ir_para_topo_no_inicio": False,
            "rolar_ate_o_final_no_inicio": True,
            "campos": {
                "AÇÕES MACRO": {
                    "coluna_planilha": "AÇÕES MACRO",
                    "planilha_linha1": list(coord_cel_p4),
                    "sistema_campo_linha1": [735, 760],
                    "botao_adicionar": [965, 697],
                    "tipo": "caixa_texto",
                    "tem_adicionar": True,
                    "data_repeater_add": "macro-actions",
                },
                "AÇÕES MICRO": {
                    "coluna_planilha": "AÇÕES MICRO",
                    "planilha_linha1": list(coord_cel_p4),
                    "sistema_campo_linha1": [1145, 760],
                    "botao_adicionar": [1375, 697],
                    "tipo": "caixa_texto",
                    "tem_adicionar": True,
                    "data_repeater_add": "micro-actions",
                },
                "AÇÕES EXTRAS": {
                    "coluna_planilha": "AÇÕES EXTRAS",
                    "planilha_linha1": list(coord_cel_p4),
                    "sistema_campo_linha1": [1555, 760],
                    "botao_adicionar": [1788, 697],
                    "tipo": "caixa_texto",
                    "tem_adicionar": True,
                    "data_repeater_add": "extra-actions",
                },
            },
            "caixas_bloqueadas": [],
            "passos_pagina": [
                {
                    "tipo_passo": "coluna_independente",
                    "nome_passo": "AÇÕES MACRO",
                    "campos": ["AÇÕES MACRO"],
                    "data_repeater_add": "macro-actions",
                    "scroll_down_apos_adicionar": True,
                },
                {
                    "tipo_passo": "coluna_independente",
                    "nome_passo": "AÇÕES MICRO",
                    "campos": ["AÇÕES MICRO"],
                    "data_repeater_add": "micro-actions",
                    "scroll_down_apos_adicionar": True,
                },
                {
                    "tipo_passo": "coluna_independente",
                    "nome_passo": "AÇÕES EXTRAS",
                    "campos": ["AÇÕES EXTRAS"],
                    "data_repeater_add": "extra-actions",
                    "scroll_down_apos_adicionar": True,
                },
            ],
        }

    # =========================================================================
    # PÁGINA 6: '6. Monitoramento, Anexos e Revisão' — Reaproveitamento em Memória da Etapa 4
    # Bloco: "Indicadores de sucesso (KPIs)" (SEM voltar à planilha!)
    # =========================================================================
    p6 = padroes.setdefault(
        "6. Monitoramento, Anexos e Revisão",
        {"nome_pagina": "6. Monitoramento, Anexos e Revisão", "campos": {}},
    )
    if not p6.get("campos"):
        p6["nome_bloco_1"] = "6. Monitoramento — Indicadores de sucesso (KPIs) [Memória Etapa 4]"
        p6["usar_memoria_etapa_4"] = True
        p6["nome_guia_planilha"] = "4.1. Metas e Ações"
        p6["nome_guia_sistema"] = "6. Monitoramento e Revisão"
        p6["step_index_sistema"] = 5
        p6["titulo_coluna_confirmacao"] = "Indicadores de sucesso (KPIs)"
        p6["campos"] = {
            "INDICADOR": {
                "coluna_planilha": "INDICADOR-CHAVE DE DESEMPENHO (KPI)",
                "planilha_linha1": list(coord_cel_p4),
                "sistema_campo_linha1": [880, 790],
                "botao_adicionar": [1765, 675],
                "tipo": "caixa_texto",
                "tem_adicionar": True,
                "passo_y_grupo": 110,
            },
            "FÓRMULA / DESCRIÇÃO": {
                "coluna_planilha": "MÉTRICA / FÓRMULA",
                "planilha_linha1": list(coord_cel_p4),
                "sistema_campo_linha1": [1480, 790],
                "botao_adicionar": [1765, 675],
                "tipo": "caixa_texto",
                "tem_adicionar": True,
                "passo_y_grupo": 110,
            },
        }
        p6["passos_pagina"] = [
            {
                "tipo_passo": "grupo_codependente",
                "nome_passo": "Indicadores de sucesso (KPIs)",
                "titulo_card_busca": "Indicadores de sucesso (KPIs)",
                "texto_botao_adicionar": "Adicionar KPI",
                "data_repeater_add": "kpis",
                "campos": ["INDICADOR", "FÓRMULA / DESCRIÇÃO"],
                "botao_adicionar": [1765, 675],
                "passo_y_grupo": 110,
            }
        ]
    else:
        p6.setdefault("usar_memoria_etapa_4", True)


def _executar_fases_seguintes_em_sequencia(
    nome_fase_inicial: str,
    mapeamento: Dict[str, Any],
    params: Dict[str, Any],
    mapa_html: Dict[str, Dict[str, Any]],
) -> bool:
    """
    Após concluir uma fase (ex: '4. Metas e Planos Obrigatórios'), verifica na sequência oficial:
    `3. Diagnóstico -> 4. Metas e Planos Obrigatórios -> 5. Capacitação, Comunicação e Ações Específicas -> 6. Monitoramento, Anexos e Revisão`
    se as fases seguintes já possuem blocos/campos configurados.
    Para cada fase configurada:
    1. Clica em `<button type="button" class="btn btn-primary" data-step-action="next">Próxima etapa</button>`.
    2. Se for a Etapa 6 ('6. Monitoramento, Anexos e Revisão'), reaproveita os dados salvos em memória
       durante a Etapa 4 (`CACHE_MEMORIA_ETAPA_4`: `INDICADOR-CHAVE DE DESEMPENHO (KPI)` e `MÉTRICA / FÓRMULA`),
       **sem necessidade de voltar à planilha**!
       Caso contrário (ex: Etapa 5), verifica na planilha se está na guia certa (ou seleciona via Toggle '≡' + nome) e copia.
    3. Verifica no sistema se está na guia certa e preenche todos os blocos daquela fase.
    """
    _garantir_estrutura_blocos_fases_4_5_6(mapeamento, params)
    padroes_paginas = mapeamento.get("padroes_paginas") or {}
    try:
        idx_atual = SEQUENCIA_FASES_SISTEMA.index(nome_fase_inicial)
    except ValueError:
        return True

    padrao_anterior = padroes_paginas.get(nome_fase_inicial) or {}

    for idx_prox in range(idx_atual + 1, len(SEQUENCIA_FASES_SISTEMA)):
        nome_fase = SEQUENCIA_FASES_SISTEMA[idx_prox]
        padrao_fase = padroes_paginas.get(nome_fase)
        if not isinstance(padrao_fase, dict):
            continue

        campos_fase = padrao_fase.get("campos") or {}
        if not campos_fase:
            print(
                f"\nℹ️  Fase seguinte na sequência ('{nome_fase}') pronta na fila "
                "(aguardando definição dos blocos/campos no treinador)."
            )
            break

        # Avança clicando no botão 'Próxima etapa' (<button data-step-action="next">Próxima etapa</button>)
        _avancar_para_proxima_etapa_do_sistema(mapeamento, padrao_anterior)

        nome_guia_f = str(padrao_fase.get("nome_guia_planilha") or nome_fase)
        coord_cel_f = padrao_fase.get("coord_celula_planilha") or params["planilha"]
        nomes_esp_f = list(campos_fase.keys())
        step_idx_f = int(padrao_fase.get("step_index_sistema", idx_prox + 2))

        usar_memoria_etapa_4 = bool(
            padrao_fase.get("usar_memoria_etapa_4", False)
            or "MONITORAMENTO" in remover_acentos(nome_fase)
        )

        if usar_memoria_etapa_4:
            linhas_f = obter_dados_etapa_4_da_memoria(mapeamento)
            if linhas_f:
                print("\n" + "=" * 72)
                print(
                    f"🧠 [REAPROVEITAMENTO DE VARIÁVEIS EM MEMÓRIA — ETAPA 4 -> '{nome_fase}']\n"
                    f"   • Utilizando matriz de {len(linhas_f)} linha(s) salva em memória na Etapa 4 "
                    "SEM voltar à planilha!\n"
                    "   • Mapeamento De/Para ('Indicadores de sucesso (KPIs)'):\n"
                    "     - 'INDICADOR-CHAVE DE DESEMPENHO (KPI)' ➔ Campo 'Indicador'\n"
                    "     - 'MÉTRICA / FÓRMULA'                  ➔ Campo 'Fórmula / descrição'"
                )
                print("=" * 72)
            else:
                print(
                    "⚠️  Cache em memória da Etapa 4 estava vazio; lendo '4.1. Metas e Ações' uma única vez para popular a memória..."
                )
                linhas_f = garantir_guia_correta_na_planilha_e_copiar(
                    padrao_bloco_ou_pagina=padrao_fase,
                    mapeamento=mapeamento,
                    coord_celula_planilha=(int(coord_cel_f[0]), int(coord_cel_f[1])),
                    nomes_esperados=list(COLUNAS_METAS_ACOES),
                    nome_guia="4.1. Metas e Ações",
                    rotulo_etapa=f"CARGA DE MEMÓRIA PARA '{nome_fase}'",
                    linhas_atuais=None,
                    verificar_aba_aberta_na_planilha=True,
                )
                salvar_dados_etapa_4_em_memoria(linhas_f, mapeamento)
        else:
            linhas_f = garantir_guia_correta_na_planilha_e_copiar(
                padrao_bloco_ou_pagina=padrao_fase,
                mapeamento=mapeamento,
                coord_celula_planilha=(int(coord_cel_f[0]), int(coord_cel_f[1])),
                nomes_esperados=nomes_esp_f,
                nome_guia=nome_guia_f,
                rotulo_etapa=f"FASE '{nome_fase}' — BLOCO 1 ('{nome_guia_f}')",
                linhas_atuais=None,
                verificar_aba_aberta_na_planilha=True,
            )

        if not linhas_f:
            return False

        _selecionar_guia_sistema_apos_copiar_tabela(
            padrao_pagina=padrao_fase,
            mapeamento=mapeamento,
            nome_pagina_ativa=nome_fase,
            step_index_padrao=step_idx_f,
            titulo_coluna_confirmacao=str(padrao_fase.get("titulo_coluna_confirmacao") or nomes_esp_f[0]),
        )

        nome_bloco_1_f = str(padrao_fase.get("nome_bloco_1") or f"{nome_fase} — Bloco 1")
        ok_f1 = _executar_roteiro_de_bloco(
            linhas=linhas_f,
            nome_bloco=nome_bloco_1_f,
            padrao_bloco=padrao_fase,
            mapeamento=mapeamento,
            params=params,
            mapa_html=mapa_html,
            ir_para_topo_no_inicio=True,
        )
        if not ok_f1:
            return False

        ok_extras = _executar_blocos_adicionais_da_pagina(
            nome_pagina=nome_fase,
            padrao_pagina=padrao_fase,
            mapeamento=mapeamento,
            params=params,
            mapa_html=mapa_html,
            step_index_sistema=step_idx_f,
        )
        if not ok_extras:
            return False

        padrao_anterior = padrao_fase

    return True


def _executar_pagina_4_metas_e_planos(
    mapeamento: Dict[str, Any],
    params: Dict[str, Any],
    mapa_html: Dict[str, Dict[str, Any]],
    nome_pagina_p4: str,
    padrao_p4: Dict[str, Any],
    linhas_iniciais: Optional[List[Dict[str, Any]]] = None,
    forcar_clique_guia_planilha: bool = True,
    rotulo_etapa: str = "ETAPA 4 — METAS E PLANOS OBRIGATÓRIOS",
) -> bool:
    """
    Executa o fluxo sequencial dos 3 blocos da página '4. Metas e Planos Obrigatórios':
    1. BLOCO 1 (Mantido intacto): Ler guia '4.1. Metas e Ações' ➔ Armazenar em memória (`CACHE_MEMORIA_ETAPA_4`)
       para reaproveitamento na Etapa 6 ➔ Colar no Bloco 1 ('Metas e ações', 7 campos + 'Adicionar ação').
    2. BLOCO 2: Mudar para a guia '4.2. Plano Orçamentário' ➔ Copiar ➔ Colar no Bloco 2
       ('Plano orçamentário e de investimentos', 7 campos + 'Adicionar linha').
    3. BLOCO 3: Mudar para a guia '4.3. Gestão de Riscos' ➔ Copiar ➔ Colar no Bloco 3
       ('Plano de gestão de riscos', 6 campos + 'Adicionar risco').
    4. Ao terminar os 3 blocos da Fase 4, segue na sequência para as Fases 5 e 6.
    """
    _garantir_estrutura_blocos_fases_4_5_6(mapeamento, params)
    nome_guia_p4 = str(padrao_p4.get("nome_guia_planilha") or "4.1. Metas e Ações")
    coord_celula_p4 = padrao_p4.get("coord_celula_planilha") or params["planilha"]

    linhas_p4 = garantir_guia_correta_na_planilha_e_copiar(
        padrao_bloco_ou_pagina=padrao_p4,
        mapeamento=mapeamento,
        coord_celula_planilha=(int(coord_celula_p4[0]), int(coord_celula_p4[1])),
        nomes_esperados=list(COLUNAS_METAS_ACOES),
        nome_guia=nome_guia_p4,
        rotulo_etapa=rotulo_etapa,
        linhas_atuais=None if forcar_clique_guia_planilha else linhas_iniciais,
        verificar_aba_aberta_na_planilha=not forcar_clique_guia_planilha,
    )

    if not linhas_p4:
        print(f"❌ Nenhuma linha válida encontrada na guia '{nome_guia_p4}'.")
        return False

    # Armazena os dados coletados na Etapa 4 ('4.1. Metas e Ações') em memória para reaproveitar na Etapa 6!
    salvar_dados_etapa_4_em_memoria(linhas_p4, mapeamento)

    # 2º Passo obrigatório APÓS copiar a tabela: verificar/selecionar a guia '4. Metas e Planos' no sistema (.wizard-nav)
    _selecionar_guia_sistema_apos_copiar_tabela(
        padrao_pagina=padrao_p4,
        mapeamento=mapeamento,
        nome_pagina_ativa=nome_pagina_p4,
        step_index_padrao=3,
        titulo_coluna_confirmacao="Metas e ações",
    )

    nome_bloco_p4 = str(
        padrao_p4.get("nome_bloco_1")
        or f"{nome_pagina_p4} — Bloco 1 (Metas e ações)"
    )
    ok_p4 = _executar_roteiro_de_bloco(
        linhas=linhas_p4,
        nome_bloco=nome_bloco_p4,
        padrao_bloco=padrao_p4,
        mapeamento=mapeamento,
        params=params,
        mapa_html=mapa_html,
        ir_para_topo_no_inicio=True,
    )
    if not ok_p4:
        return False

    # Executa sequencialmente o Bloco 2 ('4.2. Plano Orçamentário') e o Bloco 3 ('4.3. Gestão de Riscos') da Página 4
    ok_blocos_p4 = _executar_blocos_adicionais_da_pagina(
        nome_pagina=nome_pagina_p4,
        padrao_pagina=padrao_p4,
        mapeamento=mapeamento,
        params=params,
        mapa_html=mapa_html,
        step_index_sistema=3,
    )
    if not ok_blocos_p4:
        return False

    print("\n" + "=" * 72)
    print(f"🎉 Preenchimento da página '{nome_pagina_p4}' (Blocos 1, 2 e 3) concluído do início ao fim com sucesso!")
    print(f"📊 Relatório salvo em: {ARQUIVO_LOG_PROGRESSO}")
    print("=" * 72)

    # Continua para as próximas fases da sequência oficial (5. Capacitação... -> 6. Monitoramento...)
    return _executar_fases_seguintes_em_sequencia(
        nome_fase_inicial="4. Metas e Planos Obrigatórios",
        mapeamento=mapeamento,
        params=params,
        mapa_html=mapa_html,
    )


def executar_preenchimento_completo_por_coluna(
    linhas: List[Dict[str, Any]],
    pagina_forcada: Optional[str] = None,
    apenas_bloco_2: bool = False,
    forcar_clique_guia_inicial: bool = False,
    executar_ligacao_proxima_pagina: Optional[bool] = None,
) -> bool:
    """
    Executa o passo a passo completo conectado na sequência oficial das fases do sistema:
    `3. Diagnóstico` -> `4. Metas e Planos Obrigatórios` -> `5. Capacitação, Comunicação e Ações Específicas` -> `6. Monitoramento, Anexos e Revisão`.
    Em CADA fase e bloco, segue rigorosamente a regra:
    1. PRIMEIRO verifica na planilha se já está na guia ('aba') certa. Se não estiver, seleciona a guia
       usando o Toggle de Todas as Guias ('≡') + Nome de Referência (sem depender de posição variável),
       e só depois faz a cópia (`ESC -> Ctrl+A -> Ctrl+C`).
    2. ANTES de colar, verifica no sistema se está na guia certa (`.wizard-nav` / seção esperada).
    3. Após terminar os dados daquela aba, procura a aba/bloco seguinte; ao encerrar os blocos de uma página,
       pode clicar em `<button type="button" class="btn btn-primary" data-step-action="next">Próxima etapa</button>`
       e confirmar a próxima fase no sistema.
    """
    iniciar_monitor_parada_emergencia()
    mapeamento = carregar_mapeamento()
    params = obter_parametros_calibrados(mapeamento)
    mapa_html = obter_mapa_html_campos()

    colunas_planilha = list(linhas[0].keys()) if linhas else []
    nome_pagina_ativa, padrao_pagina = identificar_padrao_pagina_pelas_colunas(
        colunas_planilha, mapeamento, pagina_forcada=pagina_forcada
    )

    tem_itens_bloco_1 = _confirmar_correspondencia_guia_planilha(
        linhas, list(COLUNAS_PADRAO_SWOT), "3.1. Análise SWOT"
    )
    tem_itens_bloco_2_ja_copiados = _confirmar_correspondencia_guia_planilha(
        linhas, ["DESCRIÇÃO", "PRIORIDADE", "ÁREA DEMANDANTE", "JUSTIFICATIVA"], "3.2. Inventário de necessidades"
    )
    tem_itens_metas_acoes = _confirmar_correspondencia_guia_planilha(
        linhas, list(COLUNAS_METAS_ACOES), "4.1. Metas e Ações"
    )

    eh_pagina_4 = (
        "METAS" in remover_acentos(nome_pagina_ativa)
        or (pagina_forcada and "METAS" in remover_acentos(pagina_forcada))
        or (tem_itens_metas_acoes and not tem_itens_bloco_1 and not tem_itens_bloco_2_ja_copiados)
    )

    if eh_pagina_4:
        padroes_paginas = mapeamento.get("padroes_paginas") or {}
        if "METAS" not in remover_acentos(nome_pagina_ativa):
            for k_pag, v_pag in padroes_paginas.items():
                if "METAS" in remover_acentos(k_pag) and isinstance(v_pag, dict):
                    nome_pagina_ativa = k_pag
                    padrao_pagina = v_pag
                    break

        return _executar_pagina_4_metas_e_planos(
            mapeamento=mapeamento,
            params=params,
            mapa_html=mapa_html,
            nome_pagina_p4=nome_pagina_ativa,
            padrao_p4=padrao_pagina,
            linhas_iniciais=None,
            forcar_clique_guia_planilha=True,
            rotulo_etapa="PASSO INICIAL — GUIA '4.1. Metas e Ações'",
        )

    # =========================================================================
    # 1ª FASE DA SEQUÊNCIA: '3. Diagnóstico' — Bloco 1 ('3.1. Análise SWOT')
    # =========================================================================
    if not apenas_bloco_2:
        coord_celula_p3 = padrao_pagina.get("coord_celula_planilha") or params["planilha"]
        nome_guia_p3 = str(padrao_pagina.get("nome_guia_planilha") or "3.1. Análise SWOT")
        linhas = garantir_guia_correta_na_planilha_e_copiar(
            padrao_bloco_ou_pagina=padrao_pagina,
            mapeamento=mapeamento,
            coord_celula_planilha=(int(coord_celula_p3[0]), int(coord_celula_p3[1])),
            nomes_esperados=list(COLUNAS_PADRAO_SWOT),
            nome_guia=nome_guia_p3,
            rotulo_etapa="1ª FASE ('3. Diagnóstico' - BLOCO 1) — GUIA '3.1. Análise SWOT'",
            linhas_atuais=None if forcar_clique_guia_inicial else (linhas if tem_itens_bloco_1 else None),
            verificar_aba_aberta_na_planilha=not forcar_clique_guia_inicial,
        )
        tem_itens_bloco_1 = _confirmar_correspondencia_guia_planilha(
            linhas, list(COLUNAS_PADRAO_SWOT), nome_guia_p3
        )

    # 2º Passo obrigatório APÓS verificar/copiar a guia da planilha:
    # Verificar no sistema se está na guia '3. Diagnóstico' (`data-step-index="2"`) e coluna 'Forças'
    if not apenas_bloco_2 and tem_itens_bloco_1:
        _selecionar_guia_sistema_apos_copiar_tabela(
            padrao_pagina=padrao_pagina,
            mapeamento=mapeamento,
            nome_pagina_ativa=nome_pagina_ativa,
            step_index_padrao=2,
            titulo_coluna_confirmacao="Forças",
        )

        nome_bloco_1 = str(padrao_pagina.get("nome_bloco_1") or f"{nome_pagina_ativa} — Bloco 1 (3.1. Análise SWOT)")
        ok_bloco_1 = _executar_roteiro_de_bloco(
            linhas=linhas,
            nome_bloco=nome_bloco_1,
            padrao_bloco=padrao_pagina,
            mapeamento=mapeamento,
            params=params,
            mapa_html=mapa_html,
            ir_para_topo_no_inicio=True,
        )
        if not ok_bloco_1:
            return False

    # =========================================================================
    # 1ª FASE DA SEQUÊNCIA: '3. Diagnóstico' — Bloco 2 ('3.2. Inventário de necessidades')
    # =========================================================================
    bloco_2 = padrao_pagina.get("bloco_2")
    if isinstance(bloco_2, dict):
        nome_bloco_2 = str(bloco_2.get("nome_bloco") or "Inventário de necessidades")
        nome_guia_2 = str(bloco_2.get("nome_guia_planilha") or "3.2. Inventário de necessidades")
        bloco_2_ativo = bool(bloco_2.get("ativo", True))

        if bloco_2_ativo:
            coord_celula_2 = bloco_2.get("coord_celula_planilha") or params["planilha"]
            campos_bloco_2 = bloco_2.get("campos") or {}
            nomes_esp_2 = (
                list(campos_bloco_2.keys())
                if campos_bloco_2
                else ["DESCRIÇÃO", "PRIORIDADE", "ÁREA DEMANDANTE", "JUSTIFICATIVA"]
            )

            linhas_bloco_2 = garantir_guia_correta_na_planilha_e_copiar(
                padrao_bloco_ou_pagina=bloco_2,
                mapeamento=mapeamento,
                coord_celula_planilha=(int(coord_celula_2[0]), int(coord_celula_2[1])),
                nomes_esperados=nomes_esp_2,
                nome_guia=nome_guia_2,
                rotulo_etapa="1ª FASE ('3. Diagnóstico' - BLOCO 2) — GUIA '3.2. Inventário de necessidades'",
                linhas_atuais=linhas if (apenas_bloco_2 and tem_itens_bloco_2_ja_copiados and not tem_itens_bloco_1) else None,
                verificar_aba_aberta_na_planilha=False,
            )

            # Verifica no sistema se está na guia '3. Diagnóstico' antes de colar o Bloco 2
            if apenas_bloco_2:
                _selecionar_guia_sistema_apos_copiar_tabela(
                    padrao_pagina=padrao_pagina,
                    mapeamento=mapeamento,
                    nome_pagina_ativa=nome_pagina_ativa,
                    step_index_padrao=2,
                    titulo_coluna_confirmacao="Forças",
                )

            if linhas_bloco_2:
                ok_bloco_2 = _executar_roteiro_de_bloco(
                    linhas=linhas_bloco_2,
                    nome_bloco=f"{nome_pagina_ativa} — Bloco 2 ({nome_bloco_2})",
                    padrao_bloco=bloco_2,
                    mapeamento=mapeamento,
                    params=params,
                    mapa_html=mapa_html,
                    ir_para_topo_no_inicio=False,
                    rolar_ate_o_final_no_inicio=True,
                )
                if not ok_bloco_2:
                    return False
        else:
            print("\n" + "-" * 72)
            print(
                f"🔗 [BLOCO 2 PAUSADO] '{nome_bloco_2}' -> Guia da planilha: '{nome_guia_2}'"
            )
            print("-" * 72)

    print("\n" + "=" * 72)
    print(f"🎉 Preenchimento da 1ª Fase ('{nome_pagina_ativa}') concluído do início ao fim com sucesso!")
    print(f"📊 Relatório salvo em: {ARQUIVO_LOG_PROGRESSO}")
    print("=" * 72)

    # =========================================================================
    # TRANSIÇÃO PARA A 2ª FASE ('4. Metas e Planos Obrigatórios') E SEGUINTES (5 e 6)
    # =========================================================================
    ligar_prox = (
        executar_ligacao_proxima_pagina
        if executar_ligacao_proxima_pagina is not None
        else bool(padrao_pagina.get("ligar_proxima_pagina", True))
    )
    nome_prox_pag = str(padrao_pagina.get("proxima_pagina") or "4. Metas e Planos Obrigatórios")
    padroes_paginas = mapeamento.get("padroes_paginas") or {}
    padrao_prox = padroes_paginas.get(nome_prox_pag)

    if not apenas_bloco_2 and ligar_prox and isinstance(padrao_prox, dict):
        print("\n" + "=" * 72)
        print(
            f"🔗 [AVANÇO DE FASE NO SISTEMA: '{nome_pagina_ativa}' -> '{nome_prox_pag}'] "
            "Clicando em 'Próxima etapa' e preparando a guia '4.1. Metas e Ações'..."
        )
        print("=" * 72)
        # Clica no botão '<button type="button" class="btn btn-primary" data-step-action="next">Próxima etapa</button>'
        _avancar_para_proxima_etapa_do_sistema(mapeamento, padrao_pagina)

        return _executar_pagina_4_metas_e_planos(
            mapeamento=mapeamento,
            params=params,
            mapa_html=mapa_html,
            nome_pagina_p4=nome_prox_pag,
            padrao_p4=padrao_prox,
            linhas_iniciais=None,
            forcar_clique_guia_planilha=True,
            rotulo_etapa=f"2ª FASE ('{nome_prox_pag}' - BLOCO 1) — GUIA '4.1. Metas e Ações'",
        )

    return True


def executar_bloco_ou_pagina_especifica(alvo: str) -> bool:
    """
    Permite testar isoladamente qualquer bloco individual ('3.1', '3.2', '4.1', '4.2', '4.3', '5.1', '5.3', '6.1')
    ou uma página específica inteira ('pagina3', 'pagina4', 'pagina5', 'pagina6') sem precisar rodar as demais fases.
    """
    iniciar_monitor_parada_emergencia()
    mapeamento = carregar_mapeamento()
    params = obter_parametros_calibrados(mapeamento)
    mapa_html = obter_mapa_html_campos()
    _garantir_estrutura_blocos_fases_4_5_6(mapeamento, params)

    padroes = mapeamento.get("padroes_paginas") or {}
    chave = alvo.strip().lower().replace(",", ".")

    if chave in ("tudo", "completo", "sequencia", "all"):
        return executar_preenchimento_completo_por_coluna(
            [],
            pagina_forcada="3. Diagnóstico",
            apenas_bloco_2=False,
            forcar_clique_guia_inicial=True,
            executar_ligacao_proxima_pagina=True,
        )

    # Mapeamento de atalhos de páginas inteiras
    if chave in ("pagina3", "p3", "3", "diagnostico"):
        return executar_preenchimento_completo_por_coluna(
            [],
            pagina_forcada="3. Diagnóstico",
            apenas_bloco_2=False,
            forcar_clique_guia_inicial=True,
            executar_ligacao_proxima_pagina=False,
        )

    if chave in ("pagina4", "p4", "4", "metas"):
        p4 = padroes.get("4. Metas e Planos Obrigatórios") or {}
        coord_cel = p4.get("coord_celula_planilha") or params["planilha"]
        linhas_41 = garantir_guia_correta_na_planilha_e_copiar(
            padrao_bloco_ou_pagina=p4,
            mapeamento=mapeamento,
            coord_celula_planilha=(int(coord_cel[0]), int(coord_cel[1])),
            nomes_esperados=list(COLUNAS_METAS_ACOES),
            nome_guia="4.1. Metas e Ações",
            rotulo_etapa="PÁGINA 4 — BLOCO 1 ('4.1. Metas e Ações')",
            linhas_atuais=None,
            verificar_aba_aberta_na_planilha=False,
        )
        if not linhas_41:
            return False
        salvar_dados_etapa_4_em_memoria(linhas_41, mapeamento)
        _selecionar_guia_sistema_apos_copiar_tabela(
            padrao_pagina=p4,
            mapeamento=mapeamento,
            nome_pagina_ativa="4. Metas e Planos Obrigatórios",
            step_index_padrao=3,
            titulo_coluna_confirmacao="Metas e ações",
        )
        if not _executar_roteiro_de_bloco(
            linhas=linhas_41,
            nome_bloco=str(p4.get("nome_bloco_1") or "4. Metas e Planos Obrigatórios — Bloco 1 (Metas e ações)"),
            padrao_bloco=p4,
            mapeamento=mapeamento,
            params=params,
            mapa_html=mapa_html,
            ir_para_topo_no_inicio=True,
        ):
            return False
        return _executar_blocos_adicionais_da_pagina(
            nome_pagina="4. Metas e Planos Obrigatórios",
            padrao_pagina=p4,
            mapeamento=mapeamento,
            params=params,
            mapa_html=mapa_html,
            step_index_sistema=3,
        )

    if chave in ("pagina5", "p5", "5", "capacitacao"):
        p5 = padroes.get("5. Capacitação, Comunicação e Ações Específicas") or {}
        coord_cel = p5.get("coord_celula_planilha") or params["planilha"]
        linhas_51 = garantir_guia_correta_na_planilha_e_copiar(
            padrao_bloco_ou_pagina=p5,
            mapeamento=mapeamento,
            coord_celula_planilha=(int(coord_cel[0]), int(coord_cel[1])),
            nomes_esperados=list(COLUNAS_CAPACITACAO_5_1),
            nome_guia="5.1. Capacitação",
            rotulo_etapa="PÁGINA 5 — PASSO A ('5.1. Capacitação')",
            linhas_atuais=None,
            verificar_aba_aberta_na_planilha=True,
        )
        if not linhas_51:
            return False
        _selecionar_guia_sistema_apos_copiar_tabela(
            padrao_pagina=p5,
            mapeamento=mapeamento,
            nome_pagina_ativa="5. Capacitação, Comunicação e Ações Específicas",
            step_index_padrao=4,
            titulo_coluna_confirmacao="Treinamentos técnicos",
        )
        if not _executar_roteiro_de_bloco(
            linhas=linhas_51,
            nome_bloco="5. Capacitação — Passo A (5.1. Capacitação)",
            padrao_bloco=p5,
            mapeamento=mapeamento,
            params=params,
            mapa_html=mapa_html,
            ir_para_topo_no_inicio=True,
        ):
            return False
        return _executar_blocos_adicionais_da_pagina(
            nome_pagina="5. Capacitação, Comunicação e Ações Específicas",
            padrao_pagina=p5,
            mapeamento=mapeamento,
            params=params,
            mapa_html=mapa_html,
            step_index_sistema=4,
        )

    if chave in ("pagina6", "p6", "6", "6.1", "monitoramento", "kpis"):
        p6 = padroes.get("6. Monitoramento, Anexos e Revisão") or {}
        coord_cel = p6.get("coord_celula_planilha") or params["planilha"]
        linhas_6 = obter_dados_etapa_4_da_memoria(mapeamento)
        if not linhas_6:
            print("⚠️  Memória da Etapa 4 ainda vazia; lendo '4.1. Metas e Ações' uma vez para carregar os KPIs...")
            linhas_6 = garantir_guia_correta_na_planilha_e_copiar(
                padrao_bloco_ou_pagina=p6,
                mapeamento=mapeamento,
                coord_celula_planilha=(int(coord_cel[0]), int(coord_cel[1])),
                nomes_esperados=list(COLUNAS_METAS_ACOES),
                nome_guia="4.1. Metas e Ações",
                rotulo_etapa="CARGA DE MEMÓRIA PARA ETAPA 6 ('4.1. Metas e Ações')",
                linhas_atuais=None,
                verificar_aba_aberta_na_planilha=True,
            )
            salvar_dados_etapa_4_em_memoria(linhas_6, mapeamento)
        if not linhas_6:
            return False
        _selecionar_guia_sistema_apos_copiar_tabela(
            padrao_pagina=p6,
            mapeamento=mapeamento,
            nome_pagina_ativa="6. Monitoramento, Anexos e Revisão",
            step_index_padrao=5,
            titulo_coluna_confirmacao="Indicadores de sucesso (KPIs)",
        )
        return _executar_roteiro_de_bloco(
            linhas=linhas_6,
            nome_bloco="6. Monitoramento — Indicadores de sucesso (KPIs) [Memória Etapa 4]",
            padrao_bloco=p6,
            mapeamento=mapeamento,
            params=params,
            mapa_html=mapa_html,
            ir_para_topo_no_inicio=True,
        )

    # Execução de blocos individuais específicos: 3.1, 3.2, 4.1, 4.2, 4.3, 5.1, 5.3
    if chave in ("3.1", "swot", "bloco1"):
        p3 = padroes.get("3. Diagnóstico") or {}
        coord_cel = p3.get("coord_celula_planilha") or params["planilha"]
        linhas_31 = garantir_guia_correta_na_planilha_e_copiar(
            padrao_bloco_ou_pagina=p3,
            mapeamento=mapeamento,
            coord_celula_planilha=(int(coord_cel[0]), int(coord_cel[1])),
            nomes_esperados=list(COLUNAS_PADRAO_SWOT),
            nome_guia=str(p3.get("nome_guia_planilha") or "3.1. SWOT"),
            rotulo_etapa="BLOCO 3.1 — '3.1. SWOT'",
            linhas_atuais=None,
            verificar_aba_aberta_na_planilha=True,
        )
        if not linhas_31:
            return False
        _selecionar_guia_sistema_apos_copiar_tabela(
            padrao_pagina=p3,
            mapeamento=mapeamento,
            nome_pagina_ativa="3. Diagnóstico",
            step_index_padrao=2,
            titulo_coluna_confirmacao="Forças",
        )
        return _executar_roteiro_de_bloco(
            linhas=linhas_31,
            nome_bloco="3. Diagnóstico — Bloco 1 (3.1. SWOT)",
            padrao_bloco=p3,
            mapeamento=mapeamento,
            params=params,
            mapa_html=mapa_html,
            ir_para_topo_no_inicio=True,
        )

    if chave in ("3.2", "inventario", "bloco2"):
        return executar_preenchimento_completo_por_coluna(
            [], pagina_forcada="3. Diagnóstico", apenas_bloco_2=True
        )

    if chave in ("4.1", "metas-acoes"):
        p4 = padroes.get("4. Metas e Planos Obrigatórios") or {}
        coord_cel = p4.get("coord_celula_planilha") or params["planilha"]
        linhas_41 = garantir_guia_correta_na_planilha_e_copiar(
            padrao_bloco_ou_pagina=p4,
            mapeamento=mapeamento,
            coord_celula_planilha=(int(coord_cel[0]), int(coord_cel[1])),
            nomes_esperados=list(COLUNAS_METAS_ACOES),
            nome_guia="4.1. Metas e Ações",
            rotulo_etapa="BLOCO 4.1 — '4.1. Metas e Ações'",
            linhas_atuais=None,
            verificar_aba_aberta_na_planilha=False,
        )
        if not linhas_41:
            return False
        salvar_dados_etapa_4_em_memoria(linhas_41, mapeamento)
        _selecionar_guia_sistema_apos_copiar_tabela(
            padrao_pagina=p4,
            mapeamento=mapeamento,
            nome_pagina_ativa="4. Metas e Planos Obrigatórios",
            step_index_padrao=3,
            titulo_coluna_confirmacao="Metas e ações",
        )
        return _executar_roteiro_de_bloco(
            linhas=linhas_41,
            nome_bloco="4. Metas e Planos Obrigatórios — Bloco 1 (Metas e ações)",
            padrao_bloco=p4,
            mapeamento=mapeamento,
            params=params,
            mapa_html=mapa_html,
            ir_para_topo_no_inicio=True,
        )

    if chave in ("4.2", "orcamento", "plano-orcamentario"):
        p4 = padroes.get("4. Metas e Planos Obrigatórios") or {}
        b2_p4 = p4.get("bloco_2") or {}
        coord_cel = b2_p4.get("coord_celula_planilha") or params["planilha"]
        linhas_42 = garantir_guia_correta_na_planilha_e_copiar(
            padrao_bloco_ou_pagina=b2_p4,
            mapeamento=mapeamento,
            coord_celula_planilha=(int(coord_cel[0]), int(coord_cel[1])),
            nomes_esperados=list(COLUNAS_PLANO_ORCAMENTARIO),
            nome_guia="4.2. Plano Orçamentário",
            rotulo_etapa="BLOCO 4.2 — '4.2. Plano Orçamentário'",
            linhas_atuais=None,
            verificar_aba_aberta_na_planilha=True,
        )
        if not linhas_42:
            return False
        _selecionar_guia_sistema_apos_copiar_tabela(
            padrao_pagina=p4,
            mapeamento=mapeamento,
            nome_pagina_ativa="4. Metas e Planos Obrigatórios",
            step_index_padrao=3,
            titulo_coluna_confirmacao="Plano orçamentário",
        )
        return _executar_roteiro_de_bloco(
            linhas=linhas_42,
            nome_bloco="4. Metas e Planos Obrigatórios — Bloco 2 (Plano orçamentário e de investimentos)",
            padrao_bloco=b2_p4,
            mapeamento=mapeamento,
            params=params,
            mapa_html=mapa_html,
            ir_para_topo_no_inicio=False,
            rolar_ate_o_final_no_inicio=True,
        )

    if chave in ("4.3", "riscos", "gestao-riscos"):
        p4 = padroes.get("4. Metas e Planos Obrigatórios") or {}
        extras_p4 = p4.get("blocos_adicionais") or []
        b3_p4 = extras_p4[0] if extras_p4 else {}
        coord_cel = b3_p4.get("coord_celula_planilha") or params["planilha"]
        linhas_43 = garantir_guia_correta_na_planilha_e_copiar(
            padrao_bloco_ou_pagina=b3_p4,
            mapeamento=mapeamento,
            coord_celula_planilha=(int(coord_cel[0]), int(coord_cel[1])),
            nomes_esperados=list(COLUNAS_GESTAO_RISCOS),
            nome_guia="4.3. Gestão de Riscos",
            rotulo_etapa="BLOCO 4.3 — '4.3. Gestão de Riscos'",
            linhas_atuais=None,
            verificar_aba_aberta_na_planilha=True,
        )
        if not linhas_43:
            return False
        _selecionar_guia_sistema_apos_copiar_tabela(
            padrao_pagina=p4,
            mapeamento=mapeamento,
            nome_pagina_ativa="4. Metas e Planos Obrigatórios",
            step_index_padrao=3,
            titulo_coluna_confirmacao="Plano de gestão de riscos",
        )
        return _executar_roteiro_de_bloco(
            linhas=linhas_43,
            nome_bloco="4. Metas e Planos Obrigatórios — Bloco 3 (Plano de gestão de riscos)",
            padrao_bloco=b3_p4,
            mapeamento=mapeamento,
            params=params,
            mapa_html=mapa_html,
            ir_para_topo_no_inicio=False,
            rolar_ate_o_final_no_inicio=True,
        )

    if chave in ("5.1", "capacitacao-5.1"):
        p5 = padroes.get("5. Capacitação, Comunicação e Ações Específicas") or {}
        coord_cel = p5.get("coord_celula_planilha") or params["planilha"]
        linhas_51 = garantir_guia_correta_na_planilha_e_copiar(
            padrao_bloco_ou_pagina=p5,
            mapeamento=mapeamento,
            coord_celula_planilha=(int(coord_cel[0]), int(coord_cel[1])),
            nomes_esperados=list(COLUNAS_CAPACITACAO_5_1),
            nome_guia="5.1. Capacitação",
            rotulo_etapa="BLOCO 5.1 — '5.1. Capacitação'",
            linhas_atuais=None,
            verificar_aba_aberta_na_planilha=True,
        )
        if not linhas_51:
            return False
        _selecionar_guia_sistema_apos_copiar_tabela(
            padrao_pagina=p5,
            mapeamento=mapeamento,
            nome_pagina_ativa="5. Capacitação, Comunicação e Ações Específicas",
            step_index_padrao=4,
            titulo_coluna_confirmacao="Treinamentos técnicos",
        )
        return _executar_roteiro_de_bloco(
            linhas=linhas_51,
            nome_bloco="5. Capacitação — Passo A (5.1. Capacitação)",
            padrao_bloco=p5,
            mapeamento=mapeamento,
            params=params,
            mapa_html=mapa_html,
            ir_para_topo_no_inicio=True,
        )

    if chave in ("5.3", "acoes-especificas"):
        p5 = padroes.get("5. Capacitação, Comunicação e Ações Específicas") or {}
        b2_p5 = p5.get("bloco_2") or {}
        coord_cel = b2_p5.get("coord_celula_planilha") or params["planilha"]
        linhas_53 = garantir_guia_correta_na_planilha_e_copiar(
            padrao_bloco_ou_pagina=b2_p5,
            mapeamento=mapeamento,
            coord_celula_planilha=(int(coord_cel[0]), int(coord_cel[1])),
            nomes_esperados=list(COLUNAS_ACOES_ESPECIFICAS_5_3),
            nome_guia="5.3. Ações Específicas",
            rotulo_etapa="BLOCO 5.3 — '5.3. Ações Específicas'",
            linhas_atuais=None,
            verificar_aba_aberta_na_planilha=True,
        )
        if not linhas_53:
            return False
        _selecionar_guia_sistema_apos_copiar_tabela(
            padrao_pagina=p5,
            mapeamento=mapeamento,
            nome_pagina_ativa="5. Capacitação, Comunicação e Ações Específicas",
            step_index_padrao=4,
            titulo_coluna_confirmacao="Ações macro",
        )
        return _executar_roteiro_de_bloco(
            linhas=linhas_53,
            nome_bloco="5. Capacitação — Passo B (5.3. Ações Específicas)",
            padrao_bloco=b2_p5,
            mapeamento=mapeamento,
            params=params,
            mapa_html=mapa_html,
            ir_para_topo_no_inicio=False,
            rolar_ate_o_final_no_inicio=True,
        )

    print(f"❌ Bloco ou página '{alvo}' não reconhecido. Use: 3.1, 3.2, 4.1, 4.2, 4.3, 5.1, 5.3 ou 6.1.")
    return False



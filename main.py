"""
main.py
Ponto de entrada da automação:
1. Vai sozinho até a planilha aberta, seleciona tudo (Ctrl+A) e copia (Ctrl+C).
2. Muda para a tela do sistema e usa o conjunto de coordenadas de parâmetro + Ctrl+F + HTML.
3. Clica 1x no campo, cola o conteúdo, clica em Adicionar para criar as próximas linhas e preenche toda a tabela.
"""

import argparse
import sys
import time
from typing import Any, Dict, List

from src.automacao import (
    _extrair_tabela_auto_detectando_bloco,
    copiar_tabela_da_planilha_automaticamente,
    definir_limite_linhas_teste,
    executar_bloco_ou_pagina_especifica,
    executar_preenchimento_completo_por_coluna,
    extrair_colunas_da_tabela,
    obter_limite_linhas_teste,
)
from src.config import (
    carregar_mapeamento,
    identificar_padrao_pagina_pelas_colunas,
    salvar_mapeamento,
)
from src.driver_windows import (
    iniciar_monitor_parada_emergencia,
    ler_do_clipboard,
    obter_posicao_mouse,
    tecla_pressionada,
)
from src.localizador_visual import obter_parametros_calibrados
from treinador import calibrar_parametros_rapido, ensinar_passo_a_passo_3_linhas


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Automação RPA PDTIC por Padrões de Página, Coordenadas Relativas, OCR e DOM."
    )
    parser.add_argument(
        "--calibrar",
        "--treinar",
        dest="calibrar",
        action="store_true",
        help="Abre o treinador interativo de caixas, grupos codependentes e páginas antes de iniciar.",
    )
    parser.add_argument(
        "--treinar-completo",
        "--calibrar-completo",
        dest="treinar_completo",
        action="store_true",
        help="Abre o treinamento guiado da versão completa conectada (Fases 3, 4, 5 e 6).",
    )
    parser.add_argument(
        "--ensinar",
        "--ensinar-3-linhas",
        dest="ensinar",
        nargs="?",
        const="menu",
        default=None,
        help="Modo Ensinar Passo a Passo com 3 linhas de amostra + bloqueio de bloco vizinho (ex: --ensinar ou --ensinar 4.1).",
    )
    parser.add_argument(
        "--bloco",
        "--executar",
        dest="bloco",
        type=str,
        default=None,
        help="Executa um bloco, página ou 'tudo': 'tudo', 'pagina3', 'pagina4', 'pagina5', 'pagina6', '3.1', '3.2', '4.1', '4.2', '4.3', '5.1', '5.3' ou '6.1'.",
    )
    parser.add_argument(
        "--bloco1",
        "--swot",
        dest="bloco1",
        action="store_true",
        help="Testa diretamente o Bloco 1 da Página 3 ('3.1. SWOT').",
    )
    parser.add_argument(
        "--bloco2",
        "--inventario",
        dest="bloco2",
        action="store_true",
        help="Testa diretamente o Bloco 2 da Página 3 ('3.2. Inventário de necessidades').",
    )
    parser.add_argument(
        "--pagina",
        type=str,
        default=None,
        help="Nome ou número da página caso deseje forçar um padrão específico.",
    )
    parser.add_argument(
        "--pagina3",
        "--diagnostico",
        dest="pagina3",
        action="store_true",
        help="Executa apenas a Página 3 ('3. Diagnóstico': Bloco 3.1 SWOT + Bloco 3.2 Inventário).",
    )
    parser.add_argument(
        "--pagina4",
        "--metas-acoes",
        dest="pagina4",
        action="store_true",
        help="Executa a Página 4 ('4. Metas e Planos Obrigatórios': Blocos 4.1, 4.2 e 4.3).",
    )
    parser.add_argument(
        "--pagina5",
        "--capacitacao",
        dest="pagina5",
        action="store_true",
        help="Executa a Página 5 ('5. Capacitação, Comunicação e Ações Específicas': Passos 5.1 e 5.3).",
    )
    parser.add_argument(
        "--pagina6",
        "--monitoramento",
        dest="pagina6",
        action="store_true",
        help="Executa a Página 6 ('6. Monitoramento, Anexos e Revisão': KPIs via memória da Etapa 4).",
    )
    parser.add_argument(
        "--tudo",
        "--sequencia",
        dest="tudo",
        action="store_true",
        help="Executa toda a sequência conectada: Página 3 -> Página 4 -> Página 5 -> Página 6.",
    )
    parser.add_argument(
        "--limite",
        "--linhas",
        "-n",
        dest="limite",
        type=int,
        default=None,
        help="Limita a quantidade de linhas/itens a serem copiados de cada bloco/coluna no teste (ex: --limite 5).",
    )
    return parser.parse_args()


def _atualizar_ponto_planilha(mapeamento: Dict[str, Any], px: int, py: int) -> None:
    params = dict(mapeamento.get("parametros_referencia") or {})
    params["planilha"] = [px, py]
    mapeamento["parametros_referencia"] = params
    mapeamento["planilha"] = [px, py]
    campos_det = mapeamento.get("campos_detalhados") or {}
    for _, info in campos_det.items():
        if isinstance(info, dict):
            info["planilha_linha1"] = [px, py]
    salvar_mapeamento(mapeamento)


def obter_dados_da_planilha_aberta(mapeamento: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    Faz Ctrl+A e Ctrl+C automaticamente na planilha aberta.
    Se a planilha estiver em outra posição, permite apontar com F8.
    """
    linhas = copiar_tabela_da_planilha_automaticamente(mapeamento)
    if linhas:
        return linhas

    texto_atual = ler_do_clipboard()
    linhas_clip = _extrair_tabela_auto_detectando_bloco(texto_atual, mapeamento)
    if linhas_clip:
        return linhas_clip

    print("\n⚠️  Não foi possível localizar a janela da planilha na coordenada anterior.")
    print("👉 Coloque o mouse em QUALQUER célula dentro da sua planilha aberta e aperte [F8]")
    print("   (O sistema vai fazer Ctrl+A e Ctrl+C sozinho e salvar essa posição!)")

    while tecla_pressionada("f8"):
        time.sleep(0.03)

    while True:
        x, y = obter_posicao_mouse()
        sys.stdout.write(f"\r   📍 Mouse atual: X={x:4d}, Y={y:4d} (Aperte [F8] na planilha)   ")
        sys.stdout.flush()

        if tecla_pressionada("esc"):
            print("\n🛑 Operação cancelada.")
            sys.exit(0)

        if tecla_pressionada("f8"):
            px, py = obter_posicao_mouse()
            print(f"\n   ✅ Posição da planilha registrada: ({px}, {py})")
            _atualizar_ponto_planilha(mapeamento, px, py)
            while tecla_pressionada("f8"):
                time.sleep(0.03)
            time.sleep(0.2)
            return copiar_tabela_da_planilha_automaticamente(
                mapeamento, coord_planilha_override=(px, py)
            )

        time.sleep(0.04)


def main() -> None:
    args = parse_args()

    mapeamento = carregar_mapeamento()
    if args.ensinar is not None:
        mapeamento = ensinar_passo_a_passo_3_linhas(
            bloco_alvo=None if args.ensinar == "menu" else args.ensinar,
            dados_existentes=mapeamento,
        )
        salvar_mapeamento(mapeamento)
        if not (args.bloco or args.bloco1 or args.bloco2 or args.pagina3 or args.pagina4 or args.pagina5 or args.pagina6 or args.tudo):
            resp_rodar = input(
                "\n▶️ Deseja executar agora o preenchimento completo com o padrão aprendido? [s/N]: "
            ).strip().lower()
            if resp_rodar not in ("s", "sim", "y"):
                return
            if args.ensinar != "menu":
                args.bloco = args.ensinar
    elif args.treinar_completo:
        mapeamento = calibrar_parametros_rapido(modo_completo=True)
    elif args.calibrar or (not mapeamento.get("parametros_referencia") and not mapeamento.get("campos_detalhados")):
        mapeamento = calibrar_parametros_rapido()

    params = obter_parametros_calibrados(mapeamento)
    ponto_planilha = params["planilha"]
    pagina_ativa = mapeamento.get("pagina_ativa", "3. Diagnóstico")
    padroes_salvos = list((mapeamento.get("padroes_paginas") or {}).keys())

    print("=" * 78)
    print("🤖 AUTOMAÇÃO COLUNA A COLUNA (PADRÕES DE PÁGINA + PROTEÇÃO DE CAIXAS)")
    print("=" * 78)
    print(f"📍 Planilha configurada em: ({ponto_planilha[0]}, {ponto_planilha[1]})")
    print(f"📄 Página ativa / Padrões salvos: {pagina_ativa} (Total: {len(padroes_salvos)} página(s))")
    campos_guia = mapeamento.get("campos_detalhados") or {}
    resumo_guias = " | ".join(
        f"{col}={tuple(info.get('sistema_campo_linha1', []))}"
        for col, info in campos_guia.items()
        if isinstance(info, dict) and info.get("sistema_campo_linha1")
    )
    if resumo_guias:
        print(f"📦 Guia das Caixas ({pagina_ativa}): {resumo_guias}")

    if args.limite is not None and args.limite > 0:
        definir_limite_linhas_teste(args.limite)
    else:
        try:
            resp_limite = input(
                "🔢 Quantas linhas de cada bloco/coluna deseja copiar neste teste? "
                "(Ex: digite 5, ou aperte [ENTER] para copiar TODAS): "
            ).strip()
            if resp_limite:
                val_lim = int(resp_limite)
                if val_lim > 0:
                    definir_limite_linhas_teste(val_lim)
        except (ValueError, EOFError):
            pass

    lim_ativo = obter_limite_linhas_teste()
    if lim_ativo:
        print(f"🎯 Limite definido para este teste: {lim_ativo} linha(s) por bloco/coluna.")
    else:
        print("🎯 Modo sem limite: copiando TODAS as linhas preenchidas de cada bloco/coluna.")

    print("⏳ Iniciando em 3 segundos...")
    print("   [F8 = marcar planilha | F9 = treinar caixas/páginas/barra suspensa | ESC ou F12 = PARAR TUDO]")

    fim_espera = time.time() + 3.0
    while time.time() < fim_espera:
        restante = int(fim_espera - time.time()) + 1
        sys.stdout.write(f"\r   Iniciando em {restante}s... ")
        sys.stdout.flush()

        if tecla_pressionada("esc") or tecla_pressionada("f12"):
            print("\n🛑 Execução cancelada pelo usuário.")
            sys.exit(0)

        if tecla_pressionada("f9"):
            print("\n")
            while tecla_pressionada("f9"):
                time.sleep(0.03)
            mapeamento = calibrar_parametros_rapido()
            break

        if tecla_pressionada("f8"):
            px, py = obter_posicao_mouse()
            print(f"\n   ✅ Nova posição da planilha capturada: ({px}, {py})")
            _atualizar_ponto_planilha(mapeamento, px, py)
            while tecla_pressionada("f8"):
                time.sleep(0.03)
            break

        time.sleep(0.05)
    print("\r   Iniciando agora!          ")
    print("🛡️  DEFESA ATIVA: Aperte [ESC] ou [F12] uma única vez em qualquer tela para PARAR TUDO na hora!")
    iniciar_monitor_parada_emergencia()

    if args.bloco:
        executar_bloco_ou_pagina_especifica(args.bloco)
        return
    if args.bloco1:
        executar_bloco_ou_pagina_especifica("3.1")
        return
    if args.bloco2:
        executar_bloco_ou_pagina_especifica("3.2")
        return
    if args.pagina3 and not args.tudo:
        executar_bloco_ou_pagina_especifica("pagina3")
        return
    if args.pagina4 and not args.tudo:
        executar_bloco_ou_pagina_especifica("pagina4")
        return
    if args.pagina5 and not args.tudo:
        executar_bloco_ou_pagina_especifica("pagina5")
        return
    if args.pagina6 and not args.tudo:
        executar_bloco_ou_pagina_especifica("pagina6")
        return

    if args.pagina4:
        pagina_forcada_alvo = "4. Metas e Planos Obrigatórios"
    elif args.pagina3 or args.tudo:
        pagina_forcada_alvo = "3. Diagnóstico"
    else:
        pagina_forcada_alvo = args.pagina

    pagina_inicial = pagina_forcada_alvo or pagina_ativa
    padroes_dict = mapeamento.get("padroes_paginas") or {}
    padrao_inicial = padroes_dict.get(pagina_inicial) or {}

    if (
        "DIAGN" in pagina_inicial.upper()
        or "METAS" in pagina_inicial.upper()
        or padrao_inicial.get("coord_guia_planilha")
    ):
        executar_preenchimento_completo_por_coluna(
            [],
            pagina_forcada=pagina_inicial,
            apenas_bloco_2=False,
            forcar_clique_guia_inicial=True,
            executar_ligacao_proxima_pagina=True if args.tudo else (False if args.pagina3 else None),
        )
        return

    linhas = obter_dados_da_planilha_aberta(mapeamento)
    if not linhas:
        print("❌ Nenhuma linha de dados encontrada após copiar a planilha.")
        sys.exit(1)

    colunas_planilha = list(linhas[0].keys()) if linhas else []
    nome_pag_det, padrao_pag = identificar_padrao_pagina_pelas_colunas(
        colunas_planilha, mapeamento, pagina_forcada=pagina_forcada_alvo
    )
    campos_pag = padrao_pag.get("campos") or mapeamento.get("campos_detalhados", {})

    colunas_extraidas = extrair_colunas_da_tabela(linhas, campos_pag)
    print(
        f"✅ Planilha lida! Padrão reconhecido: '{nome_pag_det}' — "
        f"{len(colunas_extraidas)} coluna(s) prontas para preenchimento uma a uma:"
    )
    for idx_col, (nome_col, itens_col) in enumerate(colunas_extraidas, start=1):
        valores_preview = ", ".join(f"'{v}'" for _, v in itens_col)
        print(f"   📌 Coluna {idx_col} ({nome_col}) [{len(itens_col)} valor(es)]: {valores_preview}")

    executar_preenchimento_completo_por_coluna(
        linhas,
        pagina_forcada=pagina_forcada_alvo,
        apenas_bloco_2=False,
        executar_ligacao_proxima_pagina=True if args.tudo else None,
    )


if __name__ == "__main__":
    main()


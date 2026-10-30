"""
treinador.py
Mapeador Interativo de Caixas por Página (Padrões de Página + Caixas de Texto + Barra Suspensa/Toggle + Zonas Protegidas).
Permite:
1. Escolher ou criar o Padrão de cada Página (ex: '3. Diagnóstico', '1. Identificação', '2. Fundamentação', etc.).
2. Digitar livremente o nome de qualquer caixa e guiar o mouse até ela com [F8].
3. Mapear Barras Suspensas de Opções / Toggles (ex: 'PRIORIDADE' -> Alta, Média, Baixa).
4. Mapear Caixas Protegidas/Bloqueadas na página para evitar que o robô cole texto em caixas não destinadas.
5. Passar de uma página para a próxima gravando o padrão de cada página em `dados/mapeamento_campos.json`.
"""

import sys
import time
from typing import Any, Dict, List, Optional, Tuple

from src.config import (
    ARQUIVO_MAPEAMENTO,
    PAGINAS_PADRAO_SISTEMA,
    SEQUENCIA_FASES_SISTEMA,
    carregar_mapeamento,
    salvar_mapeamento,
    sincronizar_padroes_paginas,
)
from src.driver_windows import (
    mover_e_clicar,
    obter_posicao_mouse,
    tecla_pressionada,
)
from src.automacao import COLUNAS_PADRAO_SWOT, _garantir_estrutura_blocos_fases_4_5_6
from src.localizador_visual import (
    inspecionar_caixa_vazia_no_ponto,
    obter_parametros_calibrados,
    selecionar_guia_planilha_via_toggle_e_nome,
)


def _aguardar_soltar_teclas() -> None:
    while (
        tecla_pressionada("f8")
        or tecla_pressionada("enter")
        or tecla_pressionada("esc")
    ):
        time.sleep(0.03)


def aguardar_captura_ponto_f8(
    mensagem_instrucao: str,
    coord_atual: Optional[List[int]] = None,
    permitir_sem_ponto: bool = False,
) -> Tuple[Optional[int], Optional[int], bool]:
    """
    Aguarda o usuário guiar o mouse até o ponto desejado e pressionar [F8].
    - Se `coord_atual` existir e o usuário apertar [ENTER], mantém `coord_atual`.
    - Se `permitir_sem_ponto` for True e o usuário apertar [ENTER] sem `coord_atual`, retorna (None, None, False).
    """
    print(f"\n👉 {mensagem_instrucao}")
    if coord_atual:
        print(
            f"   ([F8] = capturar onde está o mouse | [ENTER] = manter atual ({coord_atual[0]}, {coord_atual[1]}) | [ESC] = cancelar)"
        )
    elif permitir_sem_ponto:
        print(
            "   ([F8] = capturar onde está o mouse | [ENTER] = não tem / automático | [ESC] = cancelar)"
        )
    else:
        print("   ([F8] = capturar onde está o mouse | [ESC] = cancelar)")

    _aguardar_soltar_teclas()

    while True:
        x, y = obter_posicao_mouse()
        sys.stdout.write(f"\r   📍 Mouse atual: X={x:4d}, Y={y:4d}   ")
        sys.stdout.flush()

        if tecla_pressionada("esc"):
            print("\n   ⏹️  Captura cancelada.")
            _aguardar_soltar_teclas()
            return None, None, False

        if tecla_pressionada("enter"):
            _aguardar_soltar_teclas()
            time.sleep(0.15)
            if coord_atual:
                print(f"\n   ⏭️  Mantida coordenada atual: ({coord_atual[0]}, {coord_atual[1]})")
                return int(coord_atual[0]), int(coord_atual[1]), False
            if permitir_sem_ponto:
                print("\n   ⏭️  Mantido modo automático / sem botão separado.")
                return None, None, False

        if tecla_pressionada("f8"):
            x_cap, y_cap = obter_posicao_mouse()
            _aguardar_soltar_teclas()
            time.sleep(0.18)
            print(f"\n   🎯 Capturado com [F8]: ({x_cap}, {y_cap})")
            return int(x_cap), int(y_cap), True

        time.sleep(0.03)


def _exibir_resumo_pagina(nome_pagina: str, padrao_pagina: Dict[str, Any]) -> None:
    campos = padrao_pagina.get("campos") or {}
    bloqueadas = padrao_pagina.get("caixas_bloqueadas") or []
    passos = padrao_pagina.get("passos_pagina") or []
    nome_guia_plan = padrao_pagina.get("nome_guia_planilha") or "(padrão)"
    coord_guia_plan = padrao_pagina.get("coord_guia_planilha")
    coord_cel_plan = padrao_pagina.get("coord_celula_planilha")
    nome_guia_sis = padrao_pagina.get("nome_guia_sistema") or nome_pagina
    step_idx_sis = padrao_pagina.get("step_index_sistema")
    coord_guia_sis = padrao_pagina.get("coord_guia_sistema")
    prox_pag = padrao_pagina.get("proxima_pagina")
    ligar_prox = padrao_pagina.get("ligar_proxima_pagina", True)
    print("\n" + "-" * 78)
    print(f"📄 PÁGINA ATUAL: {nome_pagina}")
    print(
        f"   • Guia na Planilha (1º passo antes de Ctrl+A, Ctrl+C): '{nome_guia_plan}' "
        f"| Aba={coord_guia_plan} | Célula={coord_cel_plan}"
    )
    print(
        f"   • Guia no Sistema (.wizard-nav clicada APÓS copiar a tabela): '{nome_guia_sis}' "
        f"(data-step-index={step_idx_sis}) | Botão={coord_guia_sis or 'Auto (Ctrl+F + OpenCV)'}"
    )
    if prox_pag:
        st_lig = "ATIVA ✅" if ligar_prox else "PAUSADA ⏸️"
        print(f"   • Ligação automática para próxima página: '{prox_pag}' [{st_lig}]")
    print(f"   • Caixas mapeadas nesta página ({len(campos)}):")
    if not campos:
        print("     (nenhuma caixa mapeada ainda nesta página)")
    else:
        for idx, (nome_c, info_c) in enumerate(campos.items(), start=1):
            pos_c = info_c.get("sistema_campo_linha1")
            btn_c = info_c.get("botao_adicionar")
            tipo_c = info_c.get("tipo", "caixa_texto")
            grupo_c = info_c.get("grupo_codependente")
            if tipo_c == "barra_suspensa":
                opcoes_txt = ", ".join(info_c.get("opcoes_barra_suspensa") or ["Alta", "Média", "Baixa"])
                etiqueta = f"TOGGLE/SELECT [{opcoes_txt}]"
            elif tipo_c == "bloqueada":
                etiqueta = "PROTEGIDA / NÃO ALTERAR"
            else:
                etiqueta = "CAIXA DE TEXTO"
            if grupo_c:
                etiqueta += f" (Grupo: {grupo_c})"
            print(
                f"     {idx:2d}. {nome_c:<24} | {etiqueta:<38} | Caixa={pos_c} | Add={btn_c}"
            )
    if passos:
        print(f"   • Passo a passo programado para esta página ({len(passos)} etapa(s)):")
        for idx_p, p in enumerate(passos, start=1):
            tp = p.get("tipo_passo")
            campos_p = " + ".join(p.get("campos") or [])
            if tp == "grupo_codependente":
                btn_g = p.get("botao_adicionar")
                print(
                    f"     Passo {idx_p}: 🔗 [GRUPO CODEPENDENTE] '{p.get('nome_passo')}' -> "
                    f"Preenche juntos ({campos_p}) e clica 1x em Adicionar {btn_g}"
                )
            elif tp == "campo_unico":
                print(f"     Passo {idx_p}: 📌 [CAMPO ÚNICO] '{campos_p}' (sem botão Adicionar)")
            else:
                print(
                    f"     Passo {idx_p}: 📋 [COLUNA INDEPENDENTE] '{campos_p}' "
                    "(preenche toda a coluna clicando em Adicionar a cada item)"
                )
    if bloqueadas:
        print(f"   • Zonas extras protegidas/bloqueadas ({len(bloqueadas)}): {bloqueadas}")
    print("-" * 78)


def _escolher_pagina(dados_existentes: Dict[str, Any]) -> str:
    padroes = dados_existentes.get("padroes_paginas") or {}
    pagina_ativa = str(dados_existentes.get("pagina_ativa") or "3. Diagnóstico")
    nomes_paginas_sistema = list(PAGINAS_PADRAO_SISTEMA.keys())

    # Inclui eventuais páginas customizadas já criadas pelo usuário
    for nome_p in padroes.keys():
        if nome_p not in nomes_paginas_sistema:
            nomes_paginas_sistema.append(nome_p)

    print("\n" + "=" * 78)
    print("📚 PADRÕES DE PÁGINAS DO SISTEMA (ESCOLHA QUAL PÁGINA PROGRAMAR/TREINAR)")
    print("=" * 78)
    for idx, nome_p in enumerate(nomes_paginas_sistema, start=1):
        info_p = padroes.get(nome_p) or {}
        qtd_mapeadas = len(info_p.get("campos") or {})
        qtd_passos = len(info_p.get("passos_pagina") or [])
        marcador = "👉 [ATIVA]" if nome_p == pagina_ativa else "          "
        print(
            f"  [{idx}] {nome_p:<46} ({qtd_mapeadas} caixa(s), {qtd_passos} passo(s)) {marcador}"
        )
    print("  [N] Criar um novo padrão de página com nome personalizado")
    print("-" * 78)

    escolha = input(
        f"Escolha o número da página ou aperte [ENTER] para abrir '{pagina_ativa}': "
    ).strip()

    if not escolha:
        return pagina_ativa

    if escolha.upper() == "N":
        novo_nome = input("Digite o nome da nova página (ex: 7. Nova Página): ").strip()
        return novo_nome if novo_nome else pagina_ativa

    if escolha.isdigit():
        idx_esc = int(escolha) - 1
        if 0 <= idx_esc < len(nomes_paginas_sistema):
            return nomes_paginas_sistema[idx_esc]

    return escolha


def _configurar_opcoes_toggle_interativo(
    nome_limpo: str,
    info_existente: Dict[str, Any],
    opcoes_sugeridas: Optional[List[str]] = None,
) -> Tuple[List[str], Optional[Dict[str, List[int]]]]:
    """
    Permite ao usuário adaptar livremente as opções do Toggle / Barra Suspensa (cujo valor
    será lido da planilha na execução) e opcionalmente mapear o clique de cada opção.
    """
    opcoes_padrao = (
        info_existente.get("opcoes_barra_suspensa")
        or opcoes_sugeridas
        or ["Alta", "Média", "Baixa"]
    )
    texto_padrao = ", ".join(opcoes_padrao)
    print(
        f"   💡 O valor a selecionar no Toggle '{nome_limpo}' virá da coluna '{nome_limpo}' da planilha."
    )
    entrada_opcoes = input(
        f"   Digite as opções do Toggle na ordem em que aparecem (separadas por vírgula) [ENTER = {texto_padrao}]: "
    ).strip()
    if entrada_opcoes:
        opcoes_barra = [op.strip() for op in entrada_opcoes.split(",") if op.strip()]
    else:
        opcoes_barra = list(opcoes_padrao)

    opcoes_coords: Optional[Dict[str, List[int]]] = None
    resp_map_op = input(
        "   Como selecionar a opção? ([ENTER] = Automático pelo texto da planilha / [S] = Guiar mouse [F8] até cada opção aberta): "
    ).strip().upper()
    if resp_map_op == "S":
        opcoes_coords = {}
        print("   💡 Abra o Toggle/Barra Suspensa no navegador e aponte para cada opção:")
        for nome_op in opcoes_barra:
            ox, oy, _ = aguardar_captura_ponto_f8(
                mensagem_instrucao=f"Aponte o mouse para a opção '{nome_op}' e aperte [F8]",
                coord_atual=(info_existente.get("opcoes_coordenadas") or {}).get(nome_op),
                permitir_sem_ponto=True,
            )
            if ox is not None and oy is not None:
                opcoes_coords[nome_op] = [int(ox), int(oy)]

    return opcoes_barra, opcoes_coords


def _mapear_uma_caixa_interativa(
    nome_caixa: str,
    tipo_sugerido: str,
    tem_adicionar_sugerido: bool,
    opcoes_sugeridas: Optional[List[str]],
    campos_pagina: Dict[str, Any],
    planilha_salva: List[int],
    forcar_sem_perguntar_adicionar: bool = False,
    grupo_codependente: Optional[str] = None,
) -> Optional[Dict[str, Any]]:
    """
    Mapeia uma única caixa (Caixa de Texto, Barra Suspensa/Toggle ou Caixa Protegida)
    guiando o mouse até ela com [F8].
    """
    nome_limpo = nome_caixa.strip().upper()
    info_existente = campos_pagina.get(nome_limpo) or {}
    coord_atual = info_existente.get("sistema_campo_linha1")
    btn_atual = info_existente.get("botao_adicionar")

    print("\n" + "-" * 78)
    print(f"🎯 MAPEANDO CAMPO: '{nome_limpo}'")
    print("Qual é o tipo deste campo na página?")
    print("  [1] Caixa de texto para colar dados da planilha (ou [ENTER] para padrão)")
    print("  [2] Toggle / Barra suspensa de opções (valor lido da planilha, ex: Alta, Média, Baixa)")
    print("  [3] Caixa protegida / bloqueada (NÃO deve ser preenchida, apenas evitada)")

    padrao_num = (
        "2"
        if (tipo_sugerido == "barra_suspensa" or info_existente.get("tipo") == "barra_suspensa")
        else ("3" if tipo_sugerido == "bloqueada" else "1")
    )
    resp_tipo = input(f"Escolha o tipo [padrão={padrao_num}]: ").strip() or padrao_num

    if resp_tipo == "2":
        tipo_final = "barra_suspensa"
    elif resp_tipo == "3":
        tipo_final = "bloqueada"
    else:
        tipo_final = "caixa_texto"

    if tipo_final == "barra_suspensa":
        instrucao_ponto = f"Guie o mouse até o TOGGLE / BARRA SUSPENSA '{nome_limpo}' (1ª linha) e aperte [F8]"
    elif tipo_final == "bloqueada":
        instrucao_ponto = (
            f"Guie o mouse até a CAIXA PROTEGIDA '{nome_limpo}' (que NÃO deve ser alterada) e aperte [F8]"
        )
    else:
        instrucao_ponto = (
            f"Guie o mouse DENTRO da CAIXA DE TEXTO VAZIA INICIAL de '{nome_limpo}' e aperte [F8]"
        )

    cx, cy, _ = aguardar_captura_ponto_f8(
        mensagem_instrucao=instrucao_ponto,
        coord_atual=coord_atual,
        permitir_sem_ponto=False,
    )
    if cx is None or cy is None:
        return None

    btn_final: Optional[List[int]] = list(btn_atual) if btn_atual else None
    tem_adicionar_final = bool(info_existente.get("tem_adicionar", tem_adicionar_sugerido))

    if tipo_final == "bloqueada":
        tem_adicionar_final = False
        btn_final = None
        print(
            f"   🛡️  Caixa '{nome_limpo}' gravada em ({cx}, {cy}) como ZONA PROTEGIDA (o robô nunca colará texto nela)."
        )
    elif not forcar_sem_perguntar_adicionar:
        inspecao = inspecionar_caixa_vazia_no_ponto(cx, cy)
        btn_detectado = inspecao.get("botao_adicionar") or btn_atual
        if btn_detectado:
            print(
                f"   🔎 Botão 'Adicionar' próximo detectado automaticamente em: {tuple(btn_detectado)}"
            )
        resp_btn = input(
            "   Esta caixa usa botão 'Adicionar' para criar novas linhas? ([S] Sim / [N] Não / [F8] Apontar botão manualmente / [ENTER] Automático): "
        ).strip().upper()

        if resp_btn == "N":
            tem_adicionar_final = False
            btn_final = None
        elif resp_btn == "F8":
            bx, by, _ = aguardar_captura_ponto_f8(
                mensagem_instrucao=f"Guie o mouse até o botão 'Adicionar' de '{nome_limpo}' e aperte [F8]",
                coord_atual=list(btn_detectado) if btn_detectado else None,
                permitir_sem_ponto=True,
            )
            if bx is not None and by is not None:
                btn_final = [int(bx), int(by)]
                tem_adicionar_final = True
        elif resp_btn == "S":
            tem_adicionar_final = True
            if btn_detectado:
                btn_final = [int(btn_detectado[0]), int(btn_detectado[1])]
            else:
                bx, by, _ = aguardar_captura_ponto_f8(
                    mensagem_instrucao=f"Guie o mouse até o botão 'Adicionar' de '{nome_limpo}' e aperte [F8]",
                    coord_atual=None,
                    permitir_sem_ponto=True,
                )
                if bx is not None and by is not None:
                    btn_final = [int(bx), int(by)]
        else:
            if tem_adicionar_sugerido and btn_detectado:
                btn_final = [int(btn_detectado[0]), int(btn_detectado[1])]
                tem_adicionar_final = True
            elif not tem_adicionar_sugerido:
                btn_final = None
                tem_adicionar_final = False

    opcoes_barra: Optional[List[str]] = None
    opcoes_coords: Optional[Dict[str, List[int]]] = None

    if tipo_final == "barra_suspensa":
        opcoes_barra, opcoes_coords = _configurar_opcoes_toggle_interativo(
            nome_limpo=nome_limpo,
            info_existente=info_existente,
            opcoes_sugeridas=opcoes_sugeridas,
        )

    registro_campo: Dict[str, Any] = {
        "planilha_linha1": list(planilha_salva),
        "sistema_campo_linha1": [int(cx), int(cy)],
        "botao_adicionar": btn_final,
        "tipo": tipo_final,
        "tem_adicionar": tem_adicionar_final,
    }
    if grupo_codependente:
        registro_campo["grupo_codependente"] = grupo_codependente
    if opcoes_barra:
        registro_campo["opcoes_barra_suspensa"] = opcoes_barra
    if opcoes_coords:
        registro_campo["opcoes_coordenadas"] = opcoes_coords

    campos_pagina[nome_limpo] = registro_campo
    print(
        f"   ✅ Campo '{nome_limpo}' salvo! "
        f"Posição=({cx}, {cy}) | Tipo={tipo_final} | Botão Adicionar={btn_final}"
    )
    return registro_campo


def _treinar_grupo_codependente_interativo(
    padrao_atual: Dict[str, Any],
    sugestoes_pagina: List[Dict[str, Any]],
    planilha_salva: List[int],
) -> None:
    """
    Ensina ao robô uma LINHA / GRUPO CODEPENDENTE (ex: 3 campos sendo 1 Toggle na mesma linha
    que precisam ser preenchidos todos juntos ANTES de clicar no botão 'Adicionar' compartilhado).
    """
    campos_pagina: Dict[str, Any] = padrao_atual.setdefault("campos", {})
    passos_pagina: List[Dict[str, Any]] = padrao_atual.setdefault("passos_pagina", [])

    print("\n" + "=" * 78)
    print("🔗 ENSINAR LINHA / GRUPO CODEPENDENTE (PREENCHE VÁRIOS CAMPOS ANTES DE ADICIONAR)")
    print("=" * 78)
    print(
        "Use esta opção quando uma linha do sistema possui 2, 3 ou mais campos codependentes\n"
        "(por exemplo: 2 caixas de texto + 1 Toggle/Barra Suspensa) que precisam ser preenchidos\n"
        "JUNTOS na mesma linha antes de clicar no botão 'Adicionar' para abrir a próxima linha."
    )

    nome_grupo = (
        input(
            "\nDê um nome para este grupo/seção (ex: INVENTÁRIO DE NECESSIDADES) [ENTER = automático]: "
        )
        .strip()
        .upper()
    )

    qtd_str = input(
        "Quantos campos codependentes devem ser preenchidos juntos antes de clicar em Adicionar? [ENTER = 3]: "
    ).strip()
    qtd_campos = int(qtd_str) if qtd_str.isdigit() and int(qtd_str) >= 1 else 3

    if sugestoes_pagina:
        print("\nSugestões de campos conhecidos nesta página:")
        for idx_s, item_s in enumerate(sugestoes_pagina, start=1):
            t_label = "Toggle/Barra Suspensa" if item_s.get("tipo") == "barra_suspensa" else "Caixa de Texto"
            print(f"  [{idx_s}] {item_s['nome']:<28} ({t_label})")

    nomes_campos_grupo: List[str] = []
    for i in range(qtd_campos):
        while True:
            entrada_c = input(
                f"\n👉 [{i + 1}/{qtd_campos}] Fale/Digite o NOME do {i + 1}º campo da linha (como está na coluna da planilha, ou nº da sugestão): "
            ).strip()
            if not entrada_c:
                continue

            tipo_sug = "barra_suspensa" if (qtd_campos == 3 and i == 1) else "caixa_texto"
            opcoes_sug = None

            if entrada_c.isdigit() and sugestoes_pagina:
                idx_s = int(entrada_c) - 1
                if 0 <= idx_s < len(sugestoes_pagina):
                    item_s = sugestoes_pagina[idx_s]
                    entrada_c = item_s["nome"]
                    tipo_sug = item_s.get("tipo", tipo_sug)
                    opcoes_sug = item_s.get("opcoes_barra_suspensa")
            else:
                for item_s in sugestoes_pagina:
                    if item_s["nome"].upper() == entrada_c.upper():
                        tipo_sug = item_s.get("tipo", tipo_sug)
                        opcoes_sug = item_s.get("opcoes_barra_suspensa")
                        break

            nome_grupo_efetivo = nome_grupo or f"GRUPO_{entrada_c.strip().upper()}"
            reg = _mapear_uma_caixa_interativa(
                nome_caixa=entrada_c,
                tipo_sugerido=tipo_sug,
                tem_adicionar_sugerido=False,
                opcoes_sugeridas=opcoes_sug,
                campos_pagina=campos_pagina,
                planilha_salva=planilha_salva,
                forcar_sem_perguntar_adicionar=True,
                grupo_codependente=nome_grupo_efetivo,
            )
            if reg is not None:
                nomes_campos_grupo.append(entrada_c.strip().upper())
                if not nome_grupo:
                    nome_grupo = nome_grupo_efetivo
                break

    if not nomes_campos_grupo:
        return

    # Agora captura o botão 'Adicionar' compartilhado deste grupo codependente
    coord_primeiro = campos_pagina[nomes_campos_grupo[0]]["sistema_campo_linha1"]
    inspecao = inspecionar_caixa_vazia_no_ponto(int(coord_primeiro[0]), int(coord_primeiro[1]))
    btn_detectado = inspecao.get("botao_adicionar")

    print("\n" + "-" * 78)
    print(
        f"➕ AGORA APONTE O BOTÃO 'ADICIONAR' COMPARTILHADO DO GRUPO ({' + '.join(nomes_campos_grupo)})"
    )
    bx, by, _ = aguardar_captura_ponto_f8(
        mensagem_instrucao=(
            f"Guie o mouse até o botão 'Adicionar' que só deve ser clicado APÓS preencher ({' + '.join(nomes_campos_grupo)}) e aperte [F8]"
        ),
        coord_atual=list(btn_detectado) if btn_detectado else None,
        permitir_sem_ponto=True,
    )
    btn_grupo_final = [int(bx), int(by)] if (bx is not None and by is not None) else None

    # Pergunta opcional para calibrar a distância vertical exata para a 2ª linha do grupo
    passo_y_grupo: Optional[int] = None
    print(
        "\n📏 (Opcional) Se já houver uma 2ª linha aberta deste grupo na tela, você pode apontar o 1º campo da 2ª linha com [F8]"
    )
    _, y_linha2, capturou_l2 = aguardar_captura_ponto_f8(
        mensagem_instrucao=(
            f"Aponte o 1º campo ('{nomes_campos_grupo[0]}') da 2ª linha com [F8], ou aperte [ENTER] para cálculo automático"
        ),
        coord_atual=None,
        permitir_sem_ponto=True,
    )
    if capturou_l2 and y_linha2 is not None:
        passo_calc = abs(int(y_linha2) - int(coord_primeiro[1]))
        if passo_calc >= 40:
            passo_y_grupo = passo_calc
            print(f"   ✅ Distância vertical entre linhas do grupo calibrada: {passo_y_grupo}px")

    for c_nome in nomes_campos_grupo:
        if c_nome in campos_pagina:
            campos_pagina[c_nome]["botao_adicionar"] = btn_grupo_final
            campos_pagina[c_nome]["tem_adicionar"] = bool(btn_grupo_final)
            campos_pagina[c_nome]["grupo_codependente"] = nome_grupo
            if passo_y_grupo:
                campos_pagina[c_nome]["passo_y_grupo"] = passo_y_grupo

    # Atualiza `passos_pagina` removendo passos avulsos desses campos e inserindo o passo `grupo_codependente`
    novos_passos: List[Dict[str, Any]] = []
    inserido = False
    for p in passos_pagina:
        if not isinstance(p, dict):
            continue
        if p.get("tipo_passo") == "grupo_codependente" and str(p.get("nome_passo")).upper() == nome_grupo:
            novos_passos.append(
                {
                    "tipo_passo": "grupo_codependente",
                    "nome_passo": nome_grupo,
                    "campos": nomes_campos_grupo,
                    "botao_adicionar": btn_grupo_final,
                    "passo_y_grupo": passo_y_grupo,
                }
            )
            inserido = True
            continue
        campos_restantes = [c for c in (p.get("campos") or []) if c not in nomes_campos_grupo]
        if campos_restantes:
            p_copia = dict(p)
            p_copia["campos"] = campos_restantes
            novos_passos.append(p_copia)

    if not inserido:
        novos_passos.append(
            {
                "tipo_passo": "grupo_codependente",
                "nome_passo": nome_grupo,
                "campos": nomes_campos_grupo,
                "botao_adicionar": btn_grupo_final,
                "passo_y_grupo": passo_y_grupo,
            }
        )

    padrao_atual["passos_pagina"] = novos_passos
    print(
        f"\n✅ Grupo Codependente '{nome_grupo}' ({' + '.join(nomes_campos_grupo)}) gravado com sucesso no passo a passo da página!"
    )


def _gerenciar_roteiro_passos_interativo(padrao_atual: Dict[str, Any]) -> None:
    """Permite visualizar, reordenar ou remover passos do roteiro passo a passo da página."""
    passos: List[Dict[str, Any]] = padrao_atual.get("passos_pagina") or []
    if not passos:
        print("\n⚠️ Nenhum passo programado ainda nesta página.")
        return

    print("\n" + "-" * 78)
    print("📋 GERENCIAR ORDEM DO PASSO A PASSO DESTA PÁGINA:")
    for idx, p in enumerate(passos, start=1):
        tp = p.get("tipo_passo")
        campos_txt = " + ".join(p.get("campos") or [])
        print(f"  [{idx}] {p.get('nome_passo')} ({tp}) -> {campos_txt}")
    print("-" * 78)
    print("  [M] Mover um passo de posição (mudar a ordem de execução)")
    print("  [R] Remover um passo do roteiro")
    print("  [ENTER] Voltar")

    esc = input("Escolha uma ação [M/R/ENTER]: ").strip().upper()
    if esc == "M" and len(passos) >= 2:
        orig_s = input(f"Qual número de passo deseja mover (1 a {len(passos)})? ").strip()
        dest_s = input(f"Para qual posição deseja movê-lo (1 a {len(passos)})? ").strip()
        if orig_s.isdigit() and dest_s.isdigit():
            i_orig = int(orig_s) - 1
            i_dest = int(dest_s) - 1
            if 0 <= i_orig < len(passos) and 0 <= i_dest < len(passos):
                item_mov = passos.pop(i_orig)
                passos.insert(i_dest, item_mov)
                padrao_atual["passos_pagina"] = passos
                print("✅ Ordem do passo a passo atualizada!")
    elif esc == "R":
        rem_s = input(f"Qual número de passo deseja remover (1 a {len(passos)})? ").strip()
        if rem_s.isdigit():
            i_rem = int(rem_s) - 1
            if 0 <= i_rem < len(passos):
                removido = passos.pop(i_rem)
                for c_nome in removido.get("campos") or []:
                    if c_nome in (padrao_atual.get("campos") or {}):
                        del padrao_atual["campos"][c_nome]
                padrao_atual["passos_pagina"] = passos
                print(f"🗑️ Passo '{removido.get('nome_passo')}' removido da página.")


def _configurar_bloco_2_interativo(
    padrao_atual: Dict[str, Any],
    sugestoes_pagina: List[Dict[str, Any]],
    planilha_salva: List[int],
    dados_raiz: Optional[Dict[str, Any]] = None,
    pular_menu_e_guia: bool = False,
) -> None:
    """
    Configura a Fase 2 / Bloco 2 da mesma página do sistema ('Inventário de necessidades' na página '3. Diagnóstico'),
    incluindo:
    - A LIGAÇÃO entre o Bloco 1 e o Bloco 2 (clique na guia da planilha '3.2. Inventário de necessidades').
    - Os 4 campos na sequência exata da linha:
      1º 'DESCRIÇÃO' (planilha: 'DESCRIÇÃO SUSCINTA DE NECESSIDADE PRIORIZADA NO CICLO DO PDTIC')
      2º 'PRIORIDADE' [Toggle] (planilha: 'CLASSIFICAÇÃO DE PRIORIDADE')
      3º 'ÁREA DEMANDANTE' (planilha: 'ÁREA DEMANDANTE')
      4º 'JUSTIFICATIVA' [campo maior abaixo + scroll extra] (planilha: 'JUSTIFICATIVA')
    - E o botão final da linha: 'Adicionar necessidade' (data-repeater-add="needs").
    """
    bloco_2: Dict[str, Any] = padrao_atual.setdefault(
        "bloco_2",
        {
            "nome_bloco": "Inventário de necessidades",
            "nome_guia_planilha": "3.2. Inventário de necessidades",
            "ativo": True,
            "coord_guia_planilha": None,
            "coord_celula_planilha": list(planilha_salva),
            "campos": {},
            "caixas_bloqueadas": [],
            "passos_pagina": [],
        },
    )

    nome_b2 = bloco_2.get("nome_bloco") or "Inventário de necessidades"
    nome_guia = bloco_2.get("nome_guia_planilha") or "3.2. Inventário de necessidades"
    coord_guia_atual = bloco_2.get("coord_guia_planilha")
    coord_cel_atual = bloco_2.get("coord_celula_planilha") or list(planilha_salva)
    status_str = "ATIVO ✅" if bloco_2.get("ativo", True) else "PAUSADO / ENGATILHADO ⏸️"

    campos_b2: Dict[str, Any] = bloco_2.setdefault("campos", {})

    if not pular_menu_e_guia:
        print("\n" + "=" * 78)
        print(f"🔗 CONFIGURAR BLOCO 2 DA PÁGINA 3. DIAGNÓSTICO: '{nome_b2}'")
        print(f"   • Guia na planilha: '{nome_guia}'")
        print(f"   • Clique da guia na planilha: {coord_guia_atual}")
        print(f"   • Status atual do Bloco 2: {status_str}")
        print("   • Sequência de preenchimento de cada linha antes de 'Adicionar necessidade':")
        print("     1º Descrição       <- 'DESCRIÇÃO SUSCINTA DE NECESSIDADE PRIORIZADA NO CICLO DO PDTIC'")
        print("     2º Prioridade      <- 'CLASSIFICAÇÃO DE PRIORIDADE' [Toggle: Alta / Média / Baixa]")
        print("     3º Área demandante <- 'ÁREA DEMANDANTE' [Lista suspensa na planilha]")
        print("     4º Justificativa   <- 'JUSTIFICATIVA' [Campo maior abaixo + scroll extra]")
        print("     ➕ Botão final     -> 'Adicionar necessidade' (data-repeater-add=\"needs\")")
        print("=" * 78)
        print("  [1] Mapear com [F8] a GUIA '3.2. Inventário de necessidades' + os 4 campos + botão 'Adicionar necessidade'")
        print("  [2] Ativar / Desativar a execução automática do Bloco 2 após o último item de 'Ameaças'")
        print("  [3] Configurar um grupo customizado de campos para o Bloco 2")
        print("  [ENTER] Voltar")

        esc = input("\nEscolha uma ação [1/2/3/ENTER]: ").strip()
        if esc == "2":
            bloco_2["ativo"] = not bool(bloco_2.get("ativo", True))
            novo_st = "ATIVADO ✅" if bloco_2["ativo"] else "PAUSADO ⏸️ (foco apenas no Bloco 1)"
            print(f"   🔄 Status do Bloco 2 alterado para: {novo_st}")
            return

        if esc == "3":
            sugestoes_b2 = [
                s for s in sugestoes_pagina if s.get("nome") not in COLUNAS_PADRAO_SWOT
            ]
            _treinar_grupo_codependente_interativo(
                padrao_atual=bloco_2,
                sugestoes_pagina=sugestoes_b2 or sugestoes_pagina,
                planilha_salva=list(coord_cel_atual),
            )
            return

        if esc != "1":
            return

        _configurar_toggle_todas_guias_e_aba_interativo(
            padrao_bloco_ou_pagina=bloco_2,
            nome_guia_padrao=nome_guia,
            dados_raiz=dados_raiz,
        )

        cx_cel, cy_cel, _ = aguardar_captura_ponto_f8(
            mensagem_instrucao=(
                f"2º PASSO (CÉLULA DA TABELA): Com a guia '{bloco_2.get('nome_guia_planilha') or nome_guia}' aberta, "
                f"aponte para uma célula de dados da tabela e aperte [F8] (ou [ENTER] para manter {coord_cel_atual})"
            ),
            coord_atual=coord_cel_atual,
            permitir_sem_ponto=True,
        )
        if cx_cel is not None and cy_cel is not None:
            bloco_2["coord_celula_planilha"] = [int(cx_cel), int(cy_cel)]

    print("\n" + "-" * 78)
    print("📜 AGORA NO SISTEMA (3. Diagnóstico): Role a página até o FINAL para ver 'Inventário de necessidades'!")
    print("   Vamos apontar com [F8] os 4 campos da 1ª necessidade na ordem exata e por fim o botão 'Adicionar necessidade':")
    print("-" * 78)

    definicao_4_campos = [
        (
            "DESCRIÇÃO",
            "DESCRIÇÃO SUSCINTA DE NECESSIDADE PRIORIZADA NO CICLO DO PDTIC",
            "caixa_texto",
            None,
            0,
            [667, 615],
            "1º Campo 'Descrição' (caixa de texto à esquerda)",
        ),
        (
            "PRIORIDADE",
            "CLASSIFICAÇÃO DE PRIORIDADE",
            "barra_suspensa",
            ["Alta", "Média", "Baixa"],
            0,
            [1345, 615],
            "2º Campo 'Prioridade' (Toggle/Select ao lado direito de Descrição)",
        ),
        (
            "ÁREA DEMANDANTE",
            "ÁREA DEMANDANTE",
            "caixa_texto",
            None,
            0,
            [1645, 615],
            "3º Campo 'Área demandante' (caixa ao lado direito de Prioridade)",
        ),
        (
            "JUSTIFICATIVA",
            "JUSTIFICATIVA",
            "caixa_texto",
            None,
            -3,
            [667, 750],
            "4º Campo 'Justificativa' (caixa maior ABAIXO dos 3 primeiros)",
        ),
    ]

    for nome_c, col_plan, tipo_c, opcoes_c, scroll_c, coord_padrao, desc_instrucao in definicao_4_campos:
        info_ant = campos_b2.get(nome_c) or {}
        coord_ant = info_ant.get("sistema_campo_linha1") or coord_padrao
        cx, cy, _ = aguardar_captura_ponto_f8(
            mensagem_instrucao=(
                f"Aponte o mouse para o {desc_instrucao} e aperte [F8] "
                f"(ou [ENTER] para manter {coord_ant})"
            ),
            coord_atual=coord_ant,
            permitir_sem_ponto=True,
        )
        coord_final = [int(cx), int(cy)] if (cx is not None and cy is not None) else list(coord_ant)
        registro_c: Dict[str, Any] = {
            "coluna_planilha": col_plan,
            "planilha_linha1": list(bloco_2["coord_celula_planilha"]),
            "sistema_campo_linha1": coord_final,
            "botao_adicionar": info_ant.get("botao_adicionar") or [1736, 540],
            "tipo": tipo_c,
            "tem_adicionar": True,
        }
        if opcoes_c:
            registro_c["opcoes"] = opcoes_c
        if scroll_c != 0:
            registro_c["scroll_antes_campo"] = scroll_c
        campos_b2[nome_c] = registro_c
        print(f"   ✅ {nome_c} salvo em {coord_final}!")

    # Captura do botão compartilhado 'Adicionar necessidade'
    btn_ant = (campos_b2.get("DESCRIÇÃO") or {}).get("botao_adicionar") or [1736, 540]
    bx, by, _ = aguardar_captura_ponto_f8(
        mensagem_instrucao=(
            "FINAL DA LINHA: Aponte o mouse para o botão azul 'Adicionar necessidade' "
            "(<button ... data-repeater-add=\"needs\">Adicionar necessidade</button>) e aperte [F8] "
            f"(ou [ENTER] para manter {btn_ant})"
        ),
        coord_atual=btn_ant,
        permitir_sem_ponto=True,
    )
    coord_btn_final = [int(bx), int(by)] if (bx is not None and by is not None) else list(btn_ant)
    for nome_c in ("DESCRIÇÃO", "PRIORIDADE", "ÁREA DEMANDANTE", "JUSTIFICATIVA"):
        if nome_c in campos_b2:
            campos_b2[nome_c]["botao_adicionar"] = coord_btn_final

    bloco_2["passos_pagina"] = [
        {
            "tipo_passo": "grupo_codependente",
            "nome_passo": "Inventário de necessidades",
            "titulo_card_busca": "Inventário de necessidades",
            "texto_botao_adicionar": "Adicionar necessidade",
            "campos": [
                "DESCRIÇÃO",
                "PRIORIDADE",
                "ÁREA DEMANDANTE",
                "JUSTIFICATIVA",
            ],
            "botao_adicionar": coord_btn_final,
            "passo_y_grupo": 265,
        }
    ]

    bloco_2["ativo"] = True
    print(
        f"\n   🎉 Bloco 2 ('{nome_b2}') 100% configurado e ATIVADO! "
        f"Ele será executado automaticamente após o último item de 'Ameaças'."
    )


def _configurar_toggle_todas_guias_e_aba_interativo(
    padrao_bloco_ou_pagina: Optional[Dict[str, Any]] = None,
    nome_guia_padrao: str = "",
    dados_raiz: Optional[Dict[str, Any]] = None,
    alvo_dict: Optional[Dict[str, Any]] = None,
    nome_guia_sugerido: Optional[str] = None,
    reaproveitar_toggle_salvo: bool = False,
) -> None:
    """
    Ensina ao treinador como selecionar a guia ('aba') na planilha SEM depender de posições variáveis
    na barra horizontal inferior:
    1. Pede (ou reaproveita) o ponto fixo do botão 'Toggle de Todas as Guias' (`≡` — "Todas as páginas/guias")
       no canto inferior esquerdo da planilha (`coord_toggle_todas_guias`).
    2. Define o Nome de Referência da guia (`nome_guia_planilha`, ex: '3.1. Análise SWOT',
       '3.2. Inventário de necessidades', '4.1. Metas e Ações').
    3. Clica automaticamente no Toggle (`≡`) para abrir a lista de todas as guias, localiza a guia pelo
       nome de referência via OCR nativo do Windows (ou permite apontar o item na lista aberta com [F8])
       e seleciona a aba correta na planilha!
    """
    if padrao_bloco_ou_pagina is None and alvo_dict is not None:
        padrao_bloco_ou_pagina = alvo_dict
    if not nome_guia_padrao and nome_guia_sugerido:
        nome_guia_padrao = nome_guia_sugerido
    if padrao_bloco_ou_pagina is None:
        return

    nome_guia_atual = str(padrao_bloco_ou_pagina.get("nome_guia_planilha") or nome_guia_padrao)
    novo_nome = input(
        f"Nome de referência da guia ('aba') na planilha [ENTER = manter '{nome_guia_atual}']: "
    ).strip()
    if novo_nome:
        nome_guia_atual = novo_nome
    padrao_bloco_ou_pagina["nome_guia_planilha"] = nome_guia_atual

    coord_toggle_atual = (
        (dados_raiz.get("coord_toggle_todas_guias") if isinstance(dados_raiz, dict) else None)
        or padrao_bloco_ou_pagina.get("coord_toggle_todas_guias")
    )
    if reaproveitar_toggle_salvo and coord_toggle_atual is not None:
        tx, ty = int(coord_toggle_atual[0]), int(coord_toggle_atual[1])
        print(f"   ♻️  Reaproveitando Toggle '≡' já salvo em ({tx}, {ty})!")
    else:
        tx, ty, _ = aguardar_captura_ponto_f8(
            mensagem_instrucao=(
                "1º PASSO (TOGGLE DE TODAS AS GUIAS '≡' NA PLANILHA): Posicione o mouse sobre o botão "
                "TOGGLE '≡' (que mostra todas as guias no canto inferior esquerdo da planilha) e aperte [F8]"
            ),
            coord_atual=coord_toggle_atual,
            permitir_sem_ponto=coord_toggle_atual is not None,
        )
    if tx is not None and ty is not None:
        coord_toggle_final = [int(tx), int(ty)]
        padrao_bloco_ou_pagina["coord_toggle_todas_guias"] = coord_toggle_final
        if isinstance(dados_raiz, dict):
            dados_raiz["coord_toggle_todas_guias"] = coord_toggle_final
        print(f"   ✅ Toggle de Todas as Guias ('≡') salvo em ({tx}, {ty})!")

        # Abre o Toggle e testa a seleção automática pelo Nome de Referência (OCR) ou captura [F8]
        print(
            f"   🔄 Abrindo o Toggle '≡' e buscando a guia '{nome_guia_atual}' pelo nome de referência..."
        )
        coord_ocr = selecionar_guia_planilha_via_toggle_e_nome(
            nome_guia=nome_guia_atual,
            coord_toggle_todas_guias=(int(tx), int(ty)),
            coord_item_no_toggle=padrao_bloco_ou_pagina.get("coord_item_no_toggle"),
            coord_guia_fallback=padrao_bloco_ou_pagina.get("coord_guia_planilha"),
        )
        if coord_ocr is not None:
            padrao_bloco_ou_pagina["coord_item_no_toggle"] = [int(coord_ocr[0]), int(coord_ocr[1])]
            print(
                f"   ✅ Guia '{nome_guia_atual}' selecionada automaticamente via Toggle + Nome em {coord_ocr}!"
            )
        else:
            print(
                f"   👉 Vou abrir o Toggle '≡' novamente para você apontar o nome '{nome_guia_atual}' "
                "na lista aberta com [F8] (ou [ENTER] para pular):"
            )
            mover_e_clicar(x=int(tx), y=int(ty), cliques=1, botao="left")
            time.sleep(0.45)
            ix, iy, cap_item = aguardar_captura_ponto_f8(
                mensagem_instrucao=(
                    f"Com a lista do Toggle aberta, aponte para '{nome_guia_atual}' e aperte [F8]"
                ),
                coord_atual=padrao_bloco_ou_pagina.get("coord_item_no_toggle"),
                permitir_sem_ponto=True,
            )
            if cap_item and ix is not None and iy is not None:
                padrao_bloco_ou_pagina["coord_item_no_toggle"] = [int(ix), int(iy)]
                mover_e_clicar(x=int(ix), y=int(iy), cliques=1, botao="left")
                time.sleep(0.55)
                print(f"   ✅ Item '{nome_guia_atual}' no Toggle salvo em ({ix}, {iy}) e selecionado!")


def _configurar_guias_planilha_e_sistema_interativo(
    nome_pagina_atual: str,
    padrao_atual: Dict[str, Any],
    planilha_salva: List[int],
    dados_raiz: Optional[Dict[str, Any]] = None,
) -> None:
    """
    Configura de forma interativa:
    1. O Toggle de Todas as Guias (`≡`) + Nome de Referência da guia na PLANILHA (`nome_guia_planilha`)
       e a célula da tabela (`coord_celula_planilha`) para fazer `ESC -> Ctrl+A -> Ctrl+C`.
    2. A guia correspondente no SISTEMA dentro de `<div class="nav flex-column wizard-nav">`
       (`nome_guia_sistema` + `step_index_sistema` + `coord_guia_sistema`), verificada antes de colar.
    3. O botão `<button type="button" class="btn btn-primary" data-step-action="next">Próxima etapa</button>`
       e a ligação automática na sequência:
       `3. Diagnóstico -> 4. Metas e Planos Obrigatórios -> 5. Capacitação, Comunicação e Ações Específicas -> 6. Monitoramento, Anexos e Revisão`.
    """
    nome_upper = nome_pagina_atual.upper()
    if "METAS" in nome_upper:
        nome_guia_plan_padrao = "4.1. Metas e Ações"
        nome_guia_sis_padrao = "4. Metas e Planos"
        step_idx_padrao = 3
    elif "CAPACITACAO" in nome_upper or "CAPACITAÇÃO" in nome_upper:
        nome_guia_plan_padrao = "5.1. Capacitação"
        nome_guia_sis_padrao = "5. Capacitação e Comunicação"
        step_idx_padrao = 4
    elif "MONITORAMENTO" in nome_upper:
        nome_guia_plan_padrao = "6.1. Monitoramento"
        nome_guia_sis_padrao = "6. Monitoramento e Revisão"
        step_idx_padrao = 5
    else:
        nome_guia_plan_padrao = "3.1. Análise SWOT"
        nome_guia_sis_padrao = "3. Diagnóstico"
        step_idx_padrao = 2

    coord_cel_plan_atual = padrao_atual.get("coord_celula_planilha") or list(planilha_salva)
    nome_guia_sis_atual = str(padrao_atual.get("nome_guia_sistema") or nome_guia_sis_padrao)
    step_idx_atual = padrao_atual.get("step_index_sistema", step_idx_padrao)
    coord_guia_sis_atual = padrao_atual.get("coord_guia_sistema")

    print("\n" + "=" * 78)
    print(f"🧭 CONFIGURAR TOGGLE DA PLANILHA ('≡') + GUIA DO SISTEMA (.wizard-nav) PARA '{nome_pagina_atual}'")
    print("=" * 78)

    _configurar_toggle_todas_guias_e_aba_interativo(
        padrao_bloco_ou_pagina=padrao_atual,
        nome_guia_padrao=nome_guia_plan_padrao,
        dados_raiz=dados_raiz,
    )
    nome_guia_plan_atual = str(padrao_atual.get("nome_guia_planilha") or nome_guia_plan_padrao)

    cx_cel, cy_cel, _ = aguardar_captura_ponto_f8(
        mensagem_instrucao=(
            f"2º PASSO (CÉLULA DA TABELA NA PLANILHA): Com a guia '{nome_guia_plan_atual}' aberta, "
            f"aponte para uma célula dentro da tabela (para Ctrl+A, Ctrl+C) e aperte [F8] "
            f"(ou [ENTER] para manter {coord_cel_plan_atual})"
        ),
        coord_atual=coord_cel_plan_atual,
        permitir_sem_ponto=True,
    )
    if cx_cel is not None and cy_cel is not None:
        padrao_atual["coord_celula_planilha"] = [int(cx_cel), int(cy_cel)]
        print(f"   ✅ Célula da tabela salva em ({cx_cel}, {cy_cel})!")

    novo_nome_sis = input(
        f"\nNome do botão da etapa no sistema (.wizard-nav) [ENTER = manter '{nome_guia_sis_atual}']: "
    ).strip()
    if novo_nome_sis:
        nome_guia_sis_atual = novo_nome_sis
    padrao_atual["nome_guia_sistema"] = nome_guia_sis_atual
    padrao_atual["step_index_sistema"] = step_idx_atual

    sx, sy, _ = aguardar_captura_ponto_f8(
        mensagem_instrucao=(
            f"3º PASSO (GUIA NO SISTEMA APÓS COPIAR A TABELA): No menu lateral esquerdo `.wizard-nav`, "
            f"aponte o botão '{nome_guia_sis_atual}' (data-step-index=\"{step_idx_atual}\") e aperte [F8] "
            "(ou [ENTER] para verificar/localizar automaticamente via Ctrl+F + OpenCV)"
        ),
        coord_atual=coord_guia_sis_atual,
        permitir_sem_ponto=True,
    )
    if sx is not None and sy is not None:
        padrao_atual["coord_guia_sistema"] = [int(sx), int(sy)]
        print(f"   ✅ Guia do sistema '{nome_guia_sis_atual}' salva em ({sx}, {sy})!")

    coord_prox_atual = (
        padrao_atual.get("coord_botao_proxima_etapa")
        or (dados_raiz.get("coord_botao_proxima_etapa") if isinstance(dados_raiz, dict) else None)
    )
    px, py, cap_prox = aguardar_captura_ponto_f8(
        mensagem_instrucao=(
            "4º PASSO (OPCIONAL — BOTÃO 'PRÓXIMA ETAPA' NO SISTEMA): Se desejar apontar o botão "
            "'<button data-step-action=\"next\">Próxima etapa</button>' ao final da página, aperte [F8] "
            "(ou [ENTER] para o robô localizar automaticamente via Ctrl+F + OpenCV)"
        ),
        coord_atual=coord_prox_atual,
        permitir_sem_ponto=True,
    )
    if cap_prox and px is not None and py is not None:
        padrao_atual["coord_botao_proxima_etapa"] = [int(px), int(py)]
        if isinstance(dados_raiz, dict):
            dados_raiz["coord_botao_proxima_etapa"] = [int(px), int(py)]
        print(f"   ✅ Botão 'Próxima etapa' salvo em ({px}, {py})!")

    if "DIAGN" in nome_upper:
        lig_atual = bool(padrao_atual.get("ligar_proxima_pagina", True))
        resp_lig = input(
            f"\n🔗 Deseja manter a LIGAÇÃO automática de '{nome_pagina_atual}' para a sequência "
            f"({ ' -> '.join(SEQUENCIA_FASES_SISTEMA[1:]) }) ativada? [S/N, ENTER={'S' if lig_atual else 'N'}]: "
        ).strip().upper()
        if resp_lig in ("S", "SIM"):
            padrao_atual["ligar_proxima_pagina"] = True
            padrao_atual["proxima_pagina"] = "4. Metas e Planos Obrigatórios"
        elif resp_lig in ("N", "NAO", "NÃO"):
            padrao_atual["ligar_proxima_pagina"] = False


def _configurar_bloco_1_metas_e_acoes_interativo(
    padrao_atual: Dict[str, Any],
    planilha_salva: List[int],
    dados_raiz: Optional[Dict[str, Any]] = None,
) -> None:
    """
    Configura de ponta a ponta o Bloco 1 ('Metas e ações') da página '4. Metas e Planos Obrigatórios':
    1. Guia na planilha '4.1. Metas e Ações' via Toggle de Todas as Guias (`≡`) + Nome Referência
       + célula da tabela (`coord_celula_planilha`).
    2. Guia no sistema `<button class="nav-link" data-step-index="3">4. Metas e Planos</button>` dentro de `.wizard-nav`
       (clicada logo após copiar a tabela).
    3. Os 7 campos na sequência da esquerda para a direita:
       - Linha 1 (3 caixas): ID, OBJETIVO ESTRATÉGICO (corta 'OE.TIC.X - '), META
       - Linha 2 (3 caixas): AÇÃO PRINCIPAL, INDICADOR, PRAZO (une 'QUADRANTE' + ' ' + 'ANO')
       - Linha 3 (1 caixa maior): RESPONSÁVEL
    4. Botão compartilhado 'Adicionar ação' (`data-repeater-add="goals-actions"`).
    """
    _configurar_guias_planilha_e_sistema_interativo(
        nome_pagina_atual="4. Metas e Planos Obrigatórios",
        padrao_atual=padrao_atual,
        planilha_salva=planilha_salva,
        dados_raiz=dados_raiz,
    )

    campos_p4: Dict[str, Any] = padrao_atual.setdefault("campos", {})
    coord_cel = padrao_atual.get("coord_celula_planilha") or list(planilha_salva)

    print("\n" + "-" * 78)
    print("📋 AGORA NO SISTEMA ('4. Metas e Planos' -> Card 'Metas e ações'):")
    print("   Aponte com [F8] os 7 campos da 1ª ação na sequência (esquerda -> direita) e o botão 'Adicionar ação':")
    print("-" * 78)

    definicao_7_campos = [
        ("ID", "ID OE.TIC", [748, 687], "1º Campo 'ID' (1ª linha, esquerda)"),
        (
            "OBJETIVO ESTRATÉGICO",
            "OBJETIVO ESTRATÉGICO DE TIC",
            [1053, 685],
            "2º Campo 'Objetivo estratégico' (1ª linha, centro — corta 'OE.TIC.X - ')",
        ),
        ("META", "META", [1496, 690], "3º Campo 'Meta' (1ª linha, direita)"),
        (
            "AÇÃO PRINCIPAL",
            "RESULTADO-CHAVE (KR) ESPERADO (AÇÃO PRINCIPAL)",
            [798, 776],
            "4º Campo 'Ação principal' (2ª linha, esquerda)",
        ),
        (
            "INDICADOR",
            "INDICADOR-CHAVE DE DESEMPENHO (KPI)",
            [1156, 775],
            "5º Campo 'Indicador' (2ª linha, centro)",
        ),
        (
            "PRAZO",
            "PRAZO (QUADRANTE + ANO)",
            [1428, 781],
            "6º Campo 'Prazo' (2ª linha, direita — une 'QUADRANTE' + ' ' + 'ANO')",
        ),
        (
            "RESPONSÁVEL",
            "RESPONSÁVEL",
            [1136, 865],
            "7º Campo 'Responsável' (3ª linha abaixo, campo maior único)",
        ),
    ]

    for nome_c, col_plan, coord_padrao, desc_inst in definicao_7_campos:
        info_ant = campos_p4.get(nome_c) or {}
        coord_ant = info_ant.get("sistema_campo_linha1") or coord_padrao
        cx, cy, _ = aguardar_captura_ponto_f8(
            mensagem_instrucao=(
                f"Aponte o mouse para o {desc_inst} e aperte [F8] "
                f"(ou [ENTER] para manter {coord_ant})"
            ),
            coord_atual=coord_ant,
            permitir_sem_ponto=True,
        )
        coord_final = [int(cx), int(cy)] if (cx is not None and cy is not None) else list(coord_ant)
        campos_p4[nome_c] = {
            "coluna_planilha": col_plan,
            "planilha_linha1": list(coord_cel),
            "sistema_campo_linha1": coord_final,
            "botao_adicionar": info_ant.get("botao_adicionar") or [1769, 572],
            "tipo": "caixa_texto",
            "tem_adicionar": True,
            "passo_y_grupo": 276,
        }
        print(f"   ✅ {nome_c} salvo em {coord_final}!")

    btn_ant = (campos_p4.get("ID") or {}).get("botao_adicionar") or [1769, 572]
    bx, by, _ = aguardar_captura_ponto_f8(
        mensagem_instrucao=(
            "FINAL DOS 7 CAMPOS: Aponte o mouse para o botão 'Adicionar ação' "
            "(<button ... data-repeater-add=\"goals-actions\">Adicionar ação</button>) e aperte [F8] "
            f"(ou [ENTER] para manter {btn_ant})"
        ),
        coord_atual=btn_ant,
        permitir_sem_ponto=True,
    )
    coord_btn_final = [int(bx), int(by)] if (bx is not None and by is not None) else list(btn_ant)
    nomes_7 = [item[0] for item in definicao_7_campos]
    for nome_c in nomes_7:
        if nome_c in campos_p4:
            campos_p4[nome_c]["botao_adicionar"] = coord_btn_final

    padrao_atual["passos_pagina"] = [
        {
            "tipo_passo": "grupo_codependente",
            "nome_passo": "Metas e ações",
            "titulo_card_busca": "Metas e ações",
            "texto_botao_adicionar": "Adicionar ação",
            "ancorado_no_topo": True,
            "campos": nomes_7,
            "botao_adicionar": coord_btn_final,
            "passo_y_grupo": 276,
        }
    ]
    print("\n   🎉 Página '4. Metas e Planos Obrigatórios' — Bloco 1 ('Metas e ações') 100% configurada!")


def _treinar_versao_completa_conectada(
    dados_existentes: Dict[str, Any],
    planilha_salva: List[int],
) -> None:
    """
    Treina de ponta a ponta a VERSÃO COMPLETA CONECTADA entre as fases e seus blocos:
    - 1ª FASE ('3. Diagnóstico'):
      1. Bloco 1 ('3.1. Análise SWOT' / '3.1. SWOT'): guia na planilha + célula para Ctrl+A/Ctrl+C
         + guia '3. Diagnóstico' no sistema (.wizard-nav) + (opcional) 4 campos SWOT.
      2. Bloco 2 ('3.2. Inventário de necessidades'): guia na planilha + célula para Ctrl+A/Ctrl+C
         + (opcional) 4 campos do Inventário e botão 'Adicionar necessidade'.
    - 2ª FASE ('4. Metas e Planos Obrigatórios'):
      3. Ligação automática para o Bloco 1 ('4.1. Metas e Ações'): guia na planilha + célula para Ctrl+A/Ctrl+C
         + guia '4. Metas e Planos' (`data-step-index="3"`) no sistema (.wizard-nav)
         + (opcional) 7 campos de 'Metas e ações' e botão 'Adicionar ação'.
    """
    padroes_paginas: Dict[str, Dict[str, Any]] = dados_existentes.setdefault("padroes_paginas", {})
    padrao_p3 = padroes_paginas.setdefault("3. Diagnóstico", {"nome_pagina": "3. Diagnóstico", "campos": {}})
    padrao_p4 = padroes_paginas.setdefault(
        "4. Metas e Planos Obrigatórios",
        {"nome_pagina": "4. Metas e Planos Obrigatórios", "campos": {}},
    )

    print("\n" + "=" * 78)
    print("🚀 TREINAMENTO GUIADO DA VERSÃO COMPLETA CONECTADA (FASE 1 + LIGAÇÃO + FASE 2)")
    print("=" * 78)
    print(
        "Sequência completa que será configurada:\n"
        "  1º [FASE 1 - BLOCO 1] Guia '3.1. Análise SWOT' ('3.1. SWOT') na planilha -> Ctrl+A/Ctrl+C\n"
        "                        -> Guia '3. Diagnóstico' no sistema -> Forças, Fraquezas, Oportunidades, Ameaças.\n"
        "  2º [FASE 1 - BLOCO 2] Guia '3.2. Inventário de necessidades' na planilha -> Ctrl+A/Ctrl+C\n"
        "                        -> 'Inventário de necessidades' (4 campos + 'Adicionar necessidade').\n"
        "  3º [FASE 2 - BLOCO 1] Guia '4.1. Metas e Ações' na planilha -> Ctrl+A/Ctrl+C\n"
        "                        -> Guia '4. Metas e Planos' no sistema -> 'Metas e ações' (7 campos + 'Adicionar ação')."
    )
    print("=" * 78)

    # 1. Fase 1 - Bloco 1 ('3.1. Análise SWOT')
    print("\n🔹 [ETAPA 1 DE 3] 1ª FASE ('3. Diagnóstico') — BLOCO 1 ('3.1. Análise SWOT')")
    ja_tem_etapa1 = bool((padrao_p3.get("campos") or {}).get("FORÇAS"))
    fazer_etapa1 = True
    if ja_tem_etapa1:
        resp_pular1 = input(
            "   ✅ A Etapa 1 ('3.1. SWOT') já possui coordenadas salvas. "
            "Deseja [ENTER/P] Pular para a Etapa 2 ('3.2. Inventário') ou [R] Refazer a Etapa 1? [P/r]: "
        ).strip().upper()
        if resp_pular1 not in ("R", "REFAZER", "S", "SIM"):
            fazer_etapa1 = False
            print("   ⏭️  Mantendo Etapa 1 ('3.1. SWOT') já salva e avançando para a Etapa 2!")

    if fazer_etapa1:
        _configurar_guias_planilha_e_sistema_interativo(
            nome_pagina_atual="3. Diagnóstico",
            padrao_atual=padrao_p3,
            planilha_salva=planilha_salva,
            dados_raiz=dados_existentes,
        )
        padrao_p3["ligar_proxima_pagina"] = True
        padrao_p3["proxima_pagina"] = "4. Metas e Planos Obrigatórios"
        salvar_mapeamento(dados_existentes)

        resp_swot = input(
            "\nDeseja remarcar também os 4 campos SWOT (FORÇAS, FRAQUEZAS, OPORTUNIDADES, AMEAÇAS)? "
            "[S = 1ª linha / 3 = ensinar com 3 linhas de amostra / ENTER = N (manter atuais)]: "
        ).strip().upper()
        if resp_swot == "3":
            ensinar_passo_a_passo_3_linhas(bloco_alvo="3.1", dados_existentes=dados_existentes)
            salvar_mapeamento(dados_existentes)
        elif resp_swot in ("S", "SIM"):
            campos_p3 = padrao_p3.setdefault("campos", {})
            for nome_swot in COLUNAS_PADRAO_SWOT:
                _mapear_uma_caixa_interativa(
                    nome_caixa=nome_swot,
                    tipo_sugerido="caixa_texto",
                    tem_adicionar_sugerido=True,
                    opcoes_sugeridas=None,
                    campos_pagina=campos_p3,
                    planilha_salva=list(planilha_salva),
                )
                salvar_mapeamento(dados_existentes)

    # 2. Fase 1 - Bloco 2 ('3.2. Inventário de necessidades') via Toggle de Todas as Guias + Nome Referência
    print("\n🔹 [ETAPA 2 DE 3] 1ª FASE ('3. Diagnóstico') — BLOCO 2 ('3.2. Inventário de necessidades')")
    bloco_2 = padrao_p3.setdefault(
        "bloco_2",
        {
            "nome_bloco": "Inventário de necessidades",
            "nome_guia_planilha": "3.2. Inventário de necessidades",
            "ativo": True,
            "campos": {},
            "passos_pagina": [],
        },
    )
    _configurar_toggle_todas_guias_e_aba_interativo(
        padrao_bloco_ou_pagina=bloco_2,
        nome_guia_padrao=bloco_2.get("nome_guia_planilha") or "3.2. Inventário de necessidades",
        dados_raiz=dados_existentes,
        reaproveitar_toggle_salvo=True,
    )
    bloco_2["coord_celula_planilha"] = list(
        padrao_p3.get("coord_celula_planilha") or bloco_2.get("coord_celula_planilha") or planilha_salva
    )
    bloco_2["ativo"] = True
    salvar_mapeamento(dados_existentes)

    resp_b2 = input(
        "\nDeseja remarcar os campos do 'Inventário de necessidades' no sistema? "
        "[S = 1ª linha / 3 = ensinar com 3 linhas de amostra / ENTER = N (manter atuais)]: "
    ).strip().upper()
    if resp_b2 == "3":
        ensinar_passo_a_passo_3_linhas(bloco_alvo="3.2", dados_existentes=dados_existentes)
        salvar_mapeamento(dados_existentes)
    elif resp_b2 in ("S", "SIM"):
        sugestoes_p3 = PAGINAS_PADRAO_SISTEMA.get("3. Diagnóstico", {}).get("caixas_sugeridas") or []
        _configurar_bloco_2_interativo(
            padrao_atual=padrao_p3,
            sugestoes_pagina=sugestoes_p3,
            planilha_salva=list(planilha_salva),
            dados_raiz=dados_existentes,
            pular_menu_e_guia=True,
        )
        salvar_mapeamento(dados_existentes)

    # 3. Fase 2 - Bloco 1 ('4. Metas e Planos Obrigatórios' -> '4.1. Metas e Ações')
    print("\n🔹 [ETAPA 3 DE 3] 2ª FASE ('4. Metas e Planos Obrigatórios') — BLOCO 1 ('4.1. Metas e Ações')")
    resp_p4_completo = input(
        "Deseja mapear/revisar na Fase 2 ('4.1. Metas e Ações'):\n"
        "  [G] Apenas as GUIAS (planilha '4.1. Metas e Ações' via Toggle + sistema '4. Metas e Planos') [padrão]\n"
        "  [T] Guias + apontar 1ª linha dos 7 campos de 'Metas e ações'\n"
        "  [3] Guias + ENSINAR COM 3 LINHAS DE AMOSTRA + BLOQUEAR BLOCO VIZINHO ('4.2')\n"
        "Escolha [G/T/3, ENTER=G]: "
    ).strip().upper()
    if resp_p4_completo in ("T", "TODOS", "7", "S", "SIM"):
        _configurar_bloco_1_metas_e_acoes_interativo(
            padrao_atual=padrao_p4,
            planilha_salva=list(planilha_salva),
            dados_raiz=dados_existentes,
        )
    elif resp_p4_completo == "3":
        _configurar_guias_planilha_e_sistema_interativo(
            nome_pagina_atual="4. Metas e Planos Obrigatórios",
            padrao_atual=padrao_p4,
            planilha_salva=list(planilha_salva),
            dados_raiz=dados_existentes,
        )
        ensinar_passo_a_passo_3_linhas(bloco_alvo="4.1", dados_existentes=dados_existentes)
    else:
        _configurar_guias_planilha_e_sistema_interativo(
            nome_pagina_atual="4. Metas e Planos Obrigatórios",
            padrao_atual=padrao_p4,
            planilha_salva=list(planilha_salva),
            dados_raiz=dados_existentes,
        )

    salvar_mapeamento(dados_existentes)

    resp_outros = input(
        "\n🎓 Deseja aproveitar para ensinar com 3 linhas de amostra mais algum bloco (ex: 4.2, 4.3, 5.1, 5.3 ou 6.1)? "
        "[Digite o código ex: 4.2, S para abrir o menu, ou ENTER para finalizar]: "
    ).strip()
    while resp_outros and resp_outros.upper() not in ("N", "NAO", "NÃO", "0"):
        alvo_extra = None if resp_outros.upper() in ("S", "SIM", "MENU") else resp_outros
        ensinar_passo_a_passo_3_linhas(bloco_alvo=alvo_extra, dados_existentes=dados_existentes)
        salvar_mapeamento(dados_existentes)
        resp_outros = input(
            "\n🎓 Deseja ensinar mais algum bloco com 3 linhas de amostra? [Código / S / ENTER para finalizar]: "
        ).strip()

    dados_existentes["pagina_ativa"] = "3. Diagnóstico"
    dados_existentes = sincronizar_padroes_paginas(dados_existentes)
    salvar_mapeamento(dados_existentes)

    print("\n" + "=" * 78)
    print("🎉 VERSÃO COMPLETA CONECTADA TREINADA E SALVA COM SUCESSO!")
    print("   • Seleção de aba da planilha via Toggle de Todas as Guias (≡) + Nome Referência ativada em todas as fases!")
    print("   • Página inicial ativa: '3. Diagnóstico' (Bloco 1: '3.1. Análise SWOT' -> Bloco 2: '3.2. Inventário de necessidades')")
    print("   • Sequência automática de fases: 3. Diagnóstico -> 4. Metas e Planos Obrigatórios -> 5. Capacitação -> 6. Monitoramento")
    print("=" * 78)


def calibrar_parametros_rapido(modo_completo: bool = False) -> Dict[str, Any]:
    """
    Mapeador Completo por Padrões de Página:
    - Permite falar/digitar o nome de qualquer caixa e guiar o mouse com [F8].
    - Permite ensinar Linhas/Grupos Codependentes (ex: 3 campos sendo 1 Toggle preenchidos juntos antes de clicar em Adicionar).
    - Permite mapear todas as caixas da página (incluindo caixas protegidas e barras suspensas/toggles).
    - Separa Bloco 1 (3.1. Análise SWOT) e Bloco 2 (3.2. Inventário de necessidades + clique na guia da planilha).
    """
    dados_existentes = carregar_mapeamento()
    dados_existentes = sincronizar_padroes_paginas(dados_existentes)

    padroes_paginas: Dict[str, Dict[str, Any]] = dados_existentes.setdefault("padroes_paginas", {})
    params_ref: Dict[str, Any] = dados_existentes.get("parametros_referencia") or {}
    planilha_salva = (
        dados_existentes.get("planilha")
        or params_ref.get("planilha")
        or [-1649, 373]
    )

    if modo_completo:
        _treinar_versao_completa_conectada(dados_existentes, list(planilha_salva))
        return dados_existentes

    nome_pagina_atual = _escolher_pagina(dados_existentes)

    while True:
        dados_existentes["pagina_ativa"] = nome_pagina_atual
        padrao_atual = padroes_paginas.setdefault(
            nome_pagina_atual,
            {
                "nome_pagina": nome_pagina_atual,
                "campos": {},
                "caixas_bloqueadas": [],
                "passos_pagina": [],
            },
        )
        campos_pagina: Dict[str, Any] = padrao_atual.setdefault("campos", {})
        caixas_bloqueadas: List[Dict[str, Any]] = padrao_atual.setdefault("caixas_bloqueadas", [])
        dados_existentes = sincronizar_padroes_paginas(dados_existentes)

        _exibir_resumo_pagina(nome_pagina_atual, padrao_atual)

        sugestoes_pagina = (
            PAGINAS_PADRAO_SISTEMA.get(nome_pagina_atual, {}).get("caixas_sugeridas") or []
        )
        if sugestoes_pagina:
            nomes_sug = [
                f"{s['nome']}" + (" [Toggle/Select]" if s.get("tipo") == "barra_suspensa" else "")
                for s in sugestoes_pagina
            ]
            print(f"💡 Campos conhecidos desta página: {', '.join(nomes_sug)}")

        bloco_2_info = padrao_atual.get("bloco_2") or {}
        st_b2 = "ATIVO ✅" if bloco_2_info.get("ativo") else "ENGATILHADO / PAUSADO ⏸️"

        print("\nO que você deseja programar nesta página?")
        print("  [1] Ensinar COLUNA INDEPENDENTE / CAMPO ÚNICO do Bloco 1 (falar nome e guiar mouse com [F8])")
        print(
            "  [2] Ensinar LINHA / GRUPO CODEPENDENTE (ex: 3 ou 7 campos preenchidos juntos antes de clicar em Adicionar)"
        )
        if "METAS" in nome_pagina_atual.upper():
            print(
                "  [3] Mapear GUIADO o Bloco 1 ('Metas e ações': Guia '4.1. Metas e Ações' + Guia '.wizard-nav' + 7 campos + 'Adicionar ação')"
            )
        else:
            print("  [3] Mapear em sequência os 4 campos do Bloco 1 (FORÇAS, FRAQUEZAS, OPORTUNIDADES, AMEAÇAS)")
        print("  [4] Apontar uma caixa PROTEGIDA/BLOQUEADA com [F8] (para nunca colar nela por engano)")
        print("  [5] Ver / Reordenar / Remover etapas do Passo a Passo desta página")
        print(
            f"  [6] Configurar / Ativar BLOCO 2 ('Inventário de necessidades' -> guia '3.2. Inventário de necessidades') [{st_b2}]"
        )
        print("  [7] Ir para a PRÓXIMA PÁGINA (ou trocar de página) para gravar outro padrão")
        print(
            "  [8] Apontar/Configurar GUIA DA PLANILHA (Ctrl+A, Ctrl+C) + GUIA DO SISTEMA (.wizard-nav) + LIGAÇÃO entre páginas"
        )
        print(
            "  [9] Treinar VERSÃO COMPLETA CONECTADA (Fase 1: '3.1. SWOT' + '3.2. Inventário' -> Fase 2: '4.1. Metas e Ações')"
        )
        print(
            "  [10] 🎓 ENSINAR PASSO A PASSO COM 3 LINHAS DE AMOSTRA + BLOQUEAR BLOCO VIZINHO (Aprende o ritmo e depois faz TODAS as linhas)"
        )
        print("  [0] Salvar tudo e sair (ou aperte [ENTER])")

        opcao = input("\nEscolha uma opção [0/1/2/3/4/5/6/7/8/9/10]: ").strip()

        if opcao in ("", "0"):
            break

        if opcao == "1":
            if sugestoes_pagina:
                print("\nVocê pode digitar qualquer nome de campo OU escolher o número de uma sugestão:")
                for idx_s, item_s in enumerate(sugestoes_pagina, start=1):
                    ja_tem = "✅ já mapeado" if item_s["nome"] in campos_pagina else "⚪ não mapeado"
                    t_s = "Toggle/Select" if item_s.get("tipo") == "barra_suspensa" else "Caixa de Texto"
                    print(f"  [{idx_s}] {item_s['nome']:<28} [{t_s:<14}] ({ja_tem})")

            entrada_nome = input("\nDigite o NOME do campo (ou o número da lista acima): ").strip()
            if not entrada_nome:
                continue

            tipo_sug = "caixa_texto"
            tem_add_sug = True
            opcoes_sug = None

            if entrada_nome.isdigit() and sugestoes_pagina:
                idx_s = int(entrada_nome) - 1
                if 0 <= idx_s < len(sugestoes_pagina):
                    item_s = sugestoes_pagina[idx_s]
                    entrada_nome = item_s["nome"]
                    tipo_sug = item_s.get("tipo", "caixa_texto")
                    tem_add_sug = bool(item_s.get("tem_adicionar", True))
                    opcoes_sug = item_s.get("opcoes_barra_suspensa")
            else:
                for item_s in sugestoes_pagina:
                    if item_s["nome"].upper() == entrada_nome.upper():
                        tipo_sug = item_s.get("tipo", "caixa_texto")
                        tem_add_sug = bool(item_s.get("tem_adicionar", True))
                        opcoes_sug = item_s.get("opcoes_barra_suspensa")
                        break

            _mapear_uma_caixa_interativa(
                nome_caixa=entrada_nome,
                tipo_sugerido=tipo_sug,
                tem_adicionar_sugerido=tem_add_sug,
                opcoes_sugeridas=opcoes_sug,
                campos_pagina=campos_pagina,
                planilha_salva=list(planilha_salva),
            )
            salvar_mapeamento(dados_existentes)

        elif opcao == "2":
            _treinar_grupo_codependente_interativo(
                padrao_atual=padrao_atual,
                sugestoes_pagina=sugestoes_pagina,
                planilha_salva=list(planilha_salva),
            )
            salvar_mapeamento(dados_existentes)

        elif opcao == "3":
            if "METAS" in nome_pagina_atual.upper():
                _configurar_bloco_1_metas_e_acoes_interativo(
                    padrao_atual=padrao_atual,
                    planilha_salva=list(planilha_salva),
                    dados_raiz=dados_existentes,
                )
                salvar_mapeamento(dados_existentes)
                continue

            if "Diagnóstico" in nome_pagina_atual:
                lista_para_mapear = [
                    s for s in sugestoes_pagina if s.get("nome") in COLUNAS_PADRAO_SWOT
                ]
            else:
                lista_para_mapear = list(sugestoes_pagina)
            if not lista_para_mapear:
                lista_para_mapear = [
                    {"nome": n, "tipo": "caixa_texto", "tem_adicionar": True}
                    for n in COLUNAS_PADRAO_SWOT
                ]
            print(
                f"\n🔄 Mapeando em sequência as {len(lista_para_mapear)} caixas do Bloco 1 de '{nome_pagina_atual}':"
            )
            for item_s in lista_para_mapear:
                _mapear_uma_caixa_interativa(
                    nome_caixa=item_s["nome"],
                    tipo_sugerido=item_s.get("tipo", "caixa_texto"),
                    tem_adicionar_sugerido=bool(item_s.get("tem_adicionar", True)),
                    opcoes_sugeridas=item_s.get("opcoes_barra_suspensa"),
                    campos_pagina=campos_pagina,
                    planilha_salva=list(planilha_salva),
                    grupo_codependente=item_s.get("grupo"),
                )
                salvar_mapeamento(dados_existentes)

        elif opcao == "4":
            nome_bloq = (
                input(
                    "Dê um nome para identificar esta caixa que NÃO deve ser alterada (ex: DESCRIÇÃO INVENTÁRIO): "
                )
                .strip()
                .upper()
                or f"CAIXA_BLOQUEADA_{len(caixas_bloqueadas) + 1}"
            )
            bx, by, _ = aguardar_captura_ponto_f8(
                mensagem_instrucao=f"Guie o mouse até a caixa '{nome_bloq}' que NÃO deve ser alterada e aperte [F8]",
                coord_atual=None,
                permitir_sem_ponto=False,
            )
            if bx is not None and by is not None:
                caixas_bloqueadas.append({"nome": nome_bloq, "coordenada": [int(bx), int(by)]})
                print(
                    f"   🛡️  Zona protegida '{nome_bloq}' adicionada em ({bx}, {by})! O robô nunca colará texto nela."
                )
                salvar_mapeamento(dados_existentes)

        elif opcao == "5":
            _gerenciar_roteiro_passos_interativo(padrao_atual)
            salvar_mapeamento(dados_existentes)

        elif opcao == "6":
            _configurar_bloco_2_interativo(
                padrao_atual=padrao_atual,
                sugestoes_pagina=sugestoes_pagina,
                planilha_salva=list(planilha_salva),
                dados_raiz=dados_existentes,
            )
            salvar_mapeamento(dados_existentes)

        elif opcao == "7":
            salvar_mapeamento(dados_existentes)
            nome_pagina_atual = _escolher_pagina(dados_existentes)

        elif opcao == "8":
            _configurar_guias_planilha_e_sistema_interativo(
                nome_pagina_atual=nome_pagina_atual,
                padrao_atual=padrao_atual,
                planilha_salva=list(planilha_salva),
                dados_raiz=dados_existentes,
            )
            salvar_mapeamento(dados_existentes)

        elif opcao == "9":
            _treinar_versao_completa_conectada(dados_existentes, list(planilha_salva))
            nome_pagina_atual = "3. Diagnóstico"

        elif opcao == "10":
            ensinar_passo_a_passo_3_linhas(
                bloco_alvo=None,
                dados_existentes=dados_existentes,
            )
            salvar_mapeamento(dados_existentes)

    # Sincroniza campos da página ativa com a raiz e salva o JSON final
    dados_existentes["pagina_ativa"] = nome_pagina_atual
    dados_existentes["planilha"] = list(planilha_salva)
    dados_existentes = sincronizar_padroes_paginas(dados_existentes)

    # Atualiza parametros_referencia caso FORÇAS esteja na página ativa
    campos_ativos = dados_existentes.get("campos_detalhados") or {}
    det_forcas = campos_ativos.get("FORÇAS") or {}
    if det_forcas.get("sistema_campo_linha1") and det_forcas.get("botao_adicionar"):
        c_forcas = det_forcas["sistema_campo_linha1"]
        b_forcas = det_forcas["botao_adicionar"]
        params_ref.setdefault("planilha", list(planilha_salva))
        params_ref["campo_ref"] = list(c_forcas)
        params_ref["botao_ref"] = list(b_forcas)
        params_ref.setdefault("ponto_neutro", [662, 200])
        params_ref["dx_botao_para_campo"] = int(c_forcas[0]) - int(b_forcas[0])
        params_ref["dy_botao_para_campo"] = int(c_forcas[1]) - int(b_forcas[1])
        params_ref.setdefault("dx_titulo_para_campo", 288)
        params_ref.setdefault("dy_titulo_para_campo", 73)
        params_ref.setdefault("dx_titulo_para_botao", 544)
        params_ref.setdefault("dy_titulo_para_botao", 0)
        params_ref["passo_y_nova_caixa"] = 76
        dados_existentes["parametros_referencia"] = params_ref

    salvar_mapeamento(dados_existentes)

    print("\n" + "=" * 78)
    print(f"🎉 PADRÕES E PASSO A PASSO DE PÁGINAS SALVOS EM: {ARQUIVO_MAPEAMENTO}")
    for nome_p, info_p in (dados_existentes.get("padroes_paginas") or {}).items():
        qtd_c = len((info_p or {}).get("campos") or {})
        qtd_passos = len((info_p or {}).get("passos_pagina") or [])
        qtd_b = len((info_p or {}).get("caixas_bloqueadas") or [])
        marc = " (PÁGINA ATIVA)" if nome_p == nome_pagina_atual else ""
        print(
            f"   • {nome_p:<45}: {qtd_c} campo(s), {qtd_passos} passo(s), {qtd_b} zona(s) protegida(s){marc}"
        )
    print("=" * 78)
    return dados_existentes


def ensinar_passo_a_passo_3_linhas(
    bloco_alvo: Optional[str] = None,
    dados_existentes: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """
    Modo 'Ensinar Passo a Passo com 3 Linhas de Amostra + Bloqueio de Bloco Vizinho':
    - O usuário demonstra 3 linhas de amostra (Linha 1 -> Adicionar -> Linha 2 -> Adicionar/Scroll -> Linha 3)
      e aponta opcionalmente o bloco vizinho de baixo com [F8] para bloqueá-lo.
    - O robô aprende:
      1. Coordenadas exatas dos campos da Linha 1 (validando se a caixa está vazia);
      2. Coordenada do botão 'Adicionar';
      3. Distância vertical real (`passo_y` / `passo_y_grupo`) entre a Linha 1 -> Linha 2 -> Linha 3;
      4. Ticks de scroll necessários a partir da 3ª linha;
      5. Barreira `y_limite_bloco_vizinho` para NUNCA escrever em um bloco vizinho abaixo.
    - Depois de ensinado com 3 linhas, na execução normal o robô preenche TODAS as linhas da planilha sozinho!
    """
    salvar_no_final = False
    if dados_existentes is None:
        dados_existentes = carregar_mapeamento()
        salvar_no_final = True

    dados_existentes = sincronizar_padroes_paginas(dados_existentes)
    params_cal = obter_parametros_calibrados(dados_existentes)
    _garantir_estrutura_blocos_fases_4_5_6(dados_existentes, params_cal)
    padroes = dados_existentes.setdefault("padroes_paginas", {})
    planilha_salva = list(
        (dados_existentes.get("parametros_referencia") or {}).get("planilha")
        or dados_existentes.get("planilha")
        or [-1534, 588]
    )

    catalogo_blocos = [
        {
            "codigo": "3.1",
            "pagina": "3. Diagnóstico",
            "chave_bloco": "raiz",
            "titulo": "3.1. Análise SWOT (Página 3 - Bloco 1: 4 Colunas Independentes)",
            "guia": "3.1. SWOT",
        },
        {
            "codigo": "3.2",
            "pagina": "3. Diagnóstico",
            "chave_bloco": "bloco_2",
            "titulo": "3.2. Inventário de necessidades (Página 3 - Bloco 2: 3 Campos por Linha)",
            "guia": "3.2. Inventário de necessidades",
        },
        {
            "codigo": "4.1",
            "pagina": "4. Metas e Planos Obrigatórios",
            "chave_bloco": "raiz",
            "titulo": "4.1. Metas e Ações (Página 4 - Bloco 1: 7 Campos por Linha)",
            "guia": "4.1. Metas e Ações",
        },
        {
            "codigo": "4.2",
            "pagina": "4. Metas e Planos Obrigatórios",
            "chave_bloco": "bloco_2",
            "titulo": "4.2. Plano Orçamentário (Página 4 - Bloco 2: 7 Campos por Linha)",
            "guia": "4.2. Plano Orçamentário",
        },
        {
            "codigo": "4.3",
            "pagina": "4. Metas e Planos Obrigatórios",
            "chave_bloco": "bloco_3",
            "titulo": "4.3. Gestão de Riscos (Página 4 - Bloco 3: 6 Campos por Linha)",
            "guia": "4.3. Gestão de Riscos",
        },
        {
            "codigo": "5.1",
            "pagina": "5. Capacitação, Comunicação e Ações Específicas",
            "chave_bloco": "raiz",
            "titulo": "5.1. Capacitação (Página 5 - Bloco 1: 2 Colunas Repeater)",
            "guia": "5.1. Capacitação",
        },
        {
            "codigo": "5.3",
            "pagina": "5. Capacitação, Comunicação e Ações Específicas",
            "chave_bloco": "bloco_2",
            "titulo": "5.3. Ações Específicas (Página 5 - Bloco 2: 3 Colunas Repeater)",
            "guia": "5.3. Ações Específicas",
        },
        {
            "codigo": "6.1",
            "pagina": "6. Monitoramento, Anexos e Revisão",
            "chave_bloco": "raiz",
            "titulo": "6.1. Indicadores de sucesso / KPIs (Página 6 - Bloco 1: 6 Campos por Linha)",
            "guia": "MEMORIA_ETAPA_4",
        },
    ]

    print("\n" + "=" * 82)
    print("🎓 MODO ENSINAR PASSO A PASSO (3 LINHAS DE AMOSTRA + BLOQUEIO DE BLOCO VIZINHO)")
    print("=" * 82)
    print("Como funciona:")
    print("  1. Você demonstra o caminho com 3 linhas de amostra (Linha 1 -> Linha 2 -> Linha 3).")
    print("  2. O robô verifica se cada caixa apontada está VAZIA e aprende a distância exata e o scroll.")
    print("  3. Você aponta o bloco vizinho de baixo com [F8] para criar uma BARREIRA INTRANSPONÍVEL.")
    print("  4. Ao rodar a automação normal depois, o robô faz TODAS as linhas da planilha sozinho!")
    print("-" * 82)

    item_escolhido: Optional[Dict[str, Any]] = None
    if bloco_alvo and bloco_alvo.strip().lower() not in ("menu", ""):
        alvo_norm = bloco_alvo.strip().upper()
        for item_cat in catalogo_blocos:
            if (
                alvo_norm == item_cat["codigo"].upper()
                or alvo_norm in item_cat["titulo"].upper()
                or alvo_norm in item_cat["guia"].upper()
            ):
                item_escolhido = item_cat
                break

    if item_escolhido is None:
        print("\nQual bloco você deseja ensinar com 3 linhas de amostra?")
        for idx_c, item_cat in enumerate(catalogo_blocos, start=1):
            print(f"  [{idx_c}] ({item_cat['codigo']}) {item_cat['titulo']}")
        print("  [0] Voltar / Cancelar")
        resp_b = input("\nEscolha o número ou código (ex: 3 ou 4.1): ").strip()
        if not resp_b or resp_b == "0":
            return dados_existentes
        if resp_b.isdigit() and 1 <= int(resp_b) <= len(catalogo_blocos):
            item_escolhido = catalogo_blocos[int(resp_b) - 1]
        else:
            for item_cat in catalogo_blocos:
                if resp_b.upper() in (item_cat["codigo"].upper(), item_cat["guia"].upper()):
                    item_escolhido = item_cat
                    break

    if item_escolhido is None:
        print("⚠️ Bloco não identificado.")
        return dados_existentes

    pagina_info = padroes.setdefault(item_escolhido["pagina"], {"nome_pagina": item_escolhido["pagina"], "campos": {}})
    chave_b = item_escolhido["chave_bloco"]
    if chave_b == "raiz":
        bloco_dict = pagina_info
    elif chave_b == "bloco_2":
        bloco_dict = pagina_info.setdefault("bloco_2", {"campos": {}, "ativo": True})
        bloco_dict["ativo"] = True
    else:
        blocos_adic = pagina_info.setdefault("blocos_adicionais", [{}])
        if not blocos_adic:
            blocos_adic.append({"campos": {}, "ativo": True})
        bloco_dict = blocos_adic[0]
        bloco_dict["ativo"] = True

    campos_bloco = bloco_dict.setdefault("campos", {})
    passos_bloco = bloco_dict.setdefault("passos_pagina", [])
    caixas_bloq = bloco_dict.setdefault("caixas_bloqueadas", [])

    print("\n" + "-" * 82)
    print(f"🎯 Bloco selecionado: {item_escolhido['titulo']}")
    print("👉 Antes de começar, deixe a tela do navegador visível no início deste bloco (com caixas vazias).")
    print("-" * 82)

    # Identifica se é grupo codependente ou colunas independentes
    passo_grupo = next((p for p in passos_bloco if p.get("tipo_passo") == "grupo_codependente"), None)

    if passo_grupo is not None:
        nomes_campos = list(passo_grupo.get("campos") or list(campos_bloco.keys()))
        print(f"\n📋 [AMOSTRA 1/3 — LINHA 1 DO BLOCO] Vamos apontar os {len(nomes_campos)} campos da 1ª linha vazia:")
        coords_l1: Dict[str, List[int]] = {}
        for idx_f, nome_f in enumerate(nomes_campos, start=1):
            reg_f = campos_bloco.setdefault(nome_f, {"coluna_planilha": nome_f, "tipo": "caixa_texto", "tem_adicionar": True})
            coord_at = reg_f.get("sistema_campo_linha1")
            cx, cy, _ = aguardar_captura_ponto_f8(
                mensagem_instrucao=f"[Linha 1 - Campo {idx_f}/{len(nomes_campos)}] Aponte a 1ª caixa VAZIA de '{nome_f}' e aperte [F8]",
                coord_atual=coord_at,
                permitir_sem_ponto=False,
            )
            if cx is None or cy is None:
                print("⏹️ Ensino interrompido.")
                return dados_existentes
            inspecionar_caixa_vazia_no_ponto(int(cx), int(cy), nome_f)
            reg_f["sistema_campo_linha1"] = [int(cx), int(cy)]
            reg_f["planilha_linha1"] = list(planilha_salva)
            coords_l1[nome_f] = [int(cx), int(cy)]

        btn_atual = passo_grupo.get("botao_adicionar")
        bx, by, _ = aguardar_captura_ponto_f8(
            mensagem_instrucao="Aponte o botão azul 'Adicionar...' deste bloco e aperte [F8]",
            coord_atual=btn_atual,
            permitir_sem_ponto=False,
        )
        if bx is not None and by is not None:
            passo_grupo["botao_adicionar"] = [int(bx), int(by)]
            for nome_f in nomes_campos:
                campos_bloco[nome_f]["botao_adicionar"] = [int(bx), int(by)]

        primeiro_campo = nomes_campos[0]
        y1_ref = int(coords_l1[primeiro_campo][1])

        print("\n" + "-" * 82)
        print("📋 [AMOSTRA 2/3 — LINHA 2 DO BLOCO]")
        resp_clicar_2 = input(
            "Deseja que o robô clique agora no botão 'Adicionar' para abrir a 2ª linha vazia? [S/n]: "
        ).strip().lower()
        if resp_clicar_2 != "n" and passo_grupo.get("botao_adicionar"):
            b_add = passo_grupo["botao_adicionar"]
            mover_e_clicar(int(b_add[0]), int(b_add[1]))
            time.sleep(0.55)

        passo_y_atual = int(passo_grupo.get("passo_y_grupo") or 215)
        coord_sug_l2 = [int(coords_l1[primeiro_campo][0]), min(860, y1_ref + passo_y_atual)]
        print(
            "   💡 Dica: Se a 2ª caixa nasceu abaixo da tela, você pode rolar um pouco a página antes de apertar [F8]"
            " (deixe a 1ª e a 2ª caixas visíveis na tela para medir a distância exata entre elas!)."
        )
        cx2, cy2, _ = aguardar_captura_ponto_f8(
            mensagem_instrucao=f"[Linha 2] Aponte o 1º campo ('{primeiro_campo}') da 2ª CAIXA VAZIA e aperte [F8]",
            coord_atual=coord_sug_l2,
            permitir_sem_ponto=False,
        )
        y2_ref = int(cy2) if cy2 is not None else (y1_ref + passo_y_atual)
        dy_1_2 = y2_ref - y1_ref
        if dy_1_2 < 85:
            # O usuário rolou a tela para ver a 2ª caixa! Mede na MESMA tela onde está o 1º campo da 1ª caixa:
            print(
                "   🔍 Como a página foi rolada, aponte agora na MESMA tela (sem rolar) o 1º campo da 1ª CAIXA acima para medir a distância exata:"
            )
            _, cy1_mesma_tela, _ = aguardar_captura_ponto_f8(
                mensagem_instrucao=f"[Referência na mesma tela] Aponte o 1º campo ('{primeiro_campo}') da 1ª CAIXA acima e aperte [F8]",
                coord_atual=[int(coords_l1[primeiro_campo][0]), max(230, y2_ref - passo_y_atual)],
                permitir_sem_ponto=False,
            )
            if cy1_mesma_tela is not None and (y2_ref - int(cy1_mesma_tela)) >= 60:
                dy_1_2 = int(y2_ref - int(cy1_mesma_tela))
            else:
                dy_1_2 = passo_y_atual
        dy_1_2 = max(45, int(dy_1_2))
        print(f"   📐 Distância exata aprendida entre caixas consecutivas (ΔY): +{dy_1_2}px!")

        print("\n" + "-" * 82)
        print("📋 [AMOSTRA 3/3 — LINHA 3 DO BLOCO + SCROLL]")
        resp_clicar_3 = input(
            "Deseja que o robô clique no botão 'Adicionar' para abrir a 3ª linha vazia? [S/n]: "
        ).strip().lower()
        if resp_clicar_3 != "n" and passo_grupo.get("botao_adicionar"):
            b_add = passo_grupo["botao_adicionar"]
            mover_e_clicar(int(b_add[0]), int(b_add[1]))
            time.sleep(0.55)

        ticks_l3_str = input(
            "Se a 3ª linha ficou baixa na tela, quantos ticks de scroll para baixo deseja aplicar antes de apontar? (0 = sem scroll, padrão = 5): "
        ).strip()
        ticks_l3 = int(ticks_l3_str) if ticks_l3_str.isdigit() else 0
        if ticks_l3 > 0:
            from src.driver_windows import rolar_pagina

            mover_e_clicar(int(coords_l1[primeiro_campo][0]), y2_ref)
            for _ in range(ticks_l3):
                rolar_pagina(-120)
                time.sleep(0.08)
            time.sleep(0.35)

        coord_sug_l3 = [int(coords_l1[primeiro_campo][0]), min(860, y2_ref + dy_1_2)]
        cx3, cy3, _ = aguardar_captura_ponto_f8(
            mensagem_instrucao=f"[Linha 3] Aponte o 1º campo ('{primeiro_campo}') da 3ª LINHA VAZIA e aperte [F8]",
            coord_atual=coord_sug_l3,
            permitir_sem_ponto=False,
        )
        y3_ref = int(cy3) if cy3 is not None else coord_sug_l3[1]

        passo_grupo["passo_y_grupo"] = int(dy_1_2)
        passo_grupo["aprendizado_3_linhas"] = {
            "linha_1_y": int(y1_ref),
            "linha_2_y": int(y2_ref),
            "linha_3_y": int(y3_ref),
            "passo_y_aprendido": int(dy_1_2),
            "ticks_scroll_apos_linha_2": int(ticks_l3),
        }
        for nome_f in nomes_campos:
            campos_bloco[nome_f]["passo_y_grupo"] = int(dy_1_2)

        print("\n" + "-" * 82)
        print("🛡️ [BARREIRA ANTI-INVASÃO DE BLOCO VIZINHO]")
        print("Se existir outro bloco logo abaixo na mesma página (ex: 'Plano orçamentário' abaixo de 'Metas e ações'),")
        print("aponte o TÍTULO ou o botão 'Adicionar' do bloco vizinho de baixo com [F8] para bloquear qualquer escrita nele!")
        coord_viz_atual = (
            [int(bx or 1200), int(passo_grupo["y_limite_bloco_vizinho"])]
            if passo_grupo.get("y_limite_bloco_vizinho")
            else None
        )
        vx, vy, _ = aguardar_captura_ponto_f8(
            mensagem_instrucao="Aponte o cabeçalho/botão do BLOCO VIZINHO DE BAIXO com [F8] (ou [ENTER] para barreira visual automática)",
            coord_atual=coord_viz_atual,
            permitir_sem_ponto=True,
        )
        if vy is not None:
            passo_grupo["y_limite_bloco_vizinho"] = int(vy)
            bloco_dict["y_limite_bloco_vizinho"] = int(vy)
            for nome_f in nomes_campos:
                campos_bloco[nome_f]["y_limite_bloco_vizinho"] = int(vy)
            caixas_bloq.append(
                {
                    "nome": f"BARREIRA_BLOCO_VIZINHO_{item_escolhido['codigo']}",
                    "coordenada": [int(vx or 1200), int(vy)],
                }
            )
            print(f"   🛡️ Barreira de bloco vizinho travada em Y={vy}! O robô NUNCA clicará em Y >= {vy}.")
        else:
            print("   🛡️ Barreira visual automática de botões azuis vizinhos mantida ativa!")

    else:
        nomes_colunas = list(campos_bloco.keys())
        print(f"\n📋 Este bloco possui {len(nomes_colunas)} coluna(s) independente(s): {', '.join(nomes_colunas)}")
        for idx_col, nome_col in enumerate(nomes_colunas, start=1):
            print(f"\n--- Coluna {idx_col}/{len(nomes_colunas)}: '{nome_col}' (Amostra de 3 Linhas) ---")
            reg_c = campos_bloco.setdefault(nome_col, {"coluna_planilha": nome_col, "tipo": "caixa_texto", "tem_adicionar": True})
            cx1, cy1, _ = aguardar_captura_ponto_f8(
                mensagem_instrucao=f"[{nome_col} - Linha 1/3] Aponte a 1ª caixa VAZIA e aperte [F8]",
                coord_atual=reg_c.get("sistema_campo_linha1"),
                permitir_sem_ponto=False,
            )
            if cx1 is None or cy1 is None:
                continue
            inspecionar_caixa_vazia_no_ponto(int(cx1), int(cy1), nome_col)
            reg_c["sistema_campo_linha1"] = [int(cx1), int(cy1)]

            bx, by, _ = aguardar_captura_ponto_f8(
                mensagem_instrucao=f"[{nome_col}] Aponte o botão 'Adicionar' desta coluna e aperte [F8]",
                coord_atual=reg_c.get("botao_adicionar"),
                permitir_sem_ponto=False,
            )
            if bx is not None and by is not None:
                reg_c["botao_adicionar"] = [int(bx), int(by)]

            passo_y_col = int(reg_c.get("passo_y") or 76)
            cx2, cy2, _ = aguardar_captura_ponto_f8(
                mensagem_instrucao=f"[{nome_col} - Linha 2/3] Clique em Adicionar (se necessário) e aponte a 2ª caixa VAZIA com [F8]",
                coord_atual=[int(cx1), int(cy1) + passo_y_col],
                permitir_sem_ponto=False,
            )
            dy_c = max(32, (int(cy2) - int(cy1))) if cy2 is not None else passo_y_col
            cx3, cy3, _ = aguardar_captura_ponto_f8(
                mensagem_instrucao=f"[{nome_col} - Linha 3/3] Clique em Adicionar (se necessário) e aponte a 3ª caixa VAZIA com [F8]",
                coord_atual=[int(cx1), int(cy2 or (cy1 + dy_c)) + dy_c],
                permitir_sem_ponto=False,
            )
            reg_c["passo_y"] = int(dy_c)
            reg_c["aprendizado_3_linhas"] = {
                "linha_1_y": int(cy1),
                "linha_2_y": int(cy2 or (cy1 + dy_c)),
                "linha_3_y": int(cy3 or (cy1 + 2 * dy_c)),
                "passo_y_aprendido": int(dy_c),
            }
            print(f"   ✅ Coluna '{nome_col}' aprendida com passo_y={dy_c}px!")

        vx, vy, _ = aguardar_captura_ponto_f8(
            mensagem_instrucao="🛡️ [Opcional] Aponte o bloco vizinho de baixo com [F8] para bloquear (ou [ENTER] para automático)",
            coord_atual=None,
            permitir_sem_ponto=True,
        )
        if vy is not None:
            bloco_dict["y_limite_bloco_vizinho"] = int(vy)
            for reg_c in campos_bloco.values():
                reg_c["y_limite_bloco_vizinho"] = int(vy)
            caixas_bloq.append(
                {
                    "nome": f"BARREIRA_BLOCO_VIZINHO_{item_escolhido['codigo']}",
                    "coordenada": [int(vx or 1200), int(vy)],
                }
            )

    dados_existentes = sincronizar_padroes_paginas(dados_existentes)
    if salvar_no_final:
        salvar_mapeamento(dados_existentes)
    print("\n" + "=" * 82)
    print(f"🎉 ENSINO DE 3 LINHAS CONCLUÍDO PARA: {item_escolhido['titulo']}")
    print("   Agora o robô já conhece o ritmo das caixas vazias e a barreira do bloco vizinho")
    print("   e fará TODAS as linhas da planilha automaticamente quando você executar!")
    print("=" * 82)
    return dados_existentes


def main() -> None:
    modo_completo_cli = "--completo" in sys.argv or "--tudo" in sys.argv
    if "--pagina4" in sys.argv or "--p4" in sys.argv:
        print("\n🎯 TREINAMENTO POR DISTÂNCIA DE CAIXAS — SOMENTE PÁGINA 4 (SEM MEXER NA PLANILHA)")
        print("  [1] 4.1. Metas e Ações (7 campos + Adicionar ação)")
        print("  [2] 4.2. Plano Orçamentário (7 campos + Adicionar linha)")
        print("  [3] 4.3. Gestão de Riscos (6 campos + Adicionar risco)")
        esc_p4 = input("Qual bloco da Página 4 deseja calibrar agora? [1/2/3, ENTER=1]: ").strip() or "1"
        mapa_p4 = {"1": "4.1", "2": "4.2", "3": "4.3", "4.1": "4.1", "4.2": "4.2", "4.3": "4.3"}
        ensinar_passo_a_passo_3_linhas(bloco_alvo=mapa_p4.get(esc_p4, "4.1"))
        return
    if "--ensinar" in sys.argv:
        idx_e = sys.argv.index("--ensinar")
        alvo_e = sys.argv[idx_e + 1] if (idx_e + 1) < len(sys.argv) and not sys.argv[idx_e + 1].startswith("-") else None
        ensinar_passo_a_passo_3_linhas(bloco_alvo=alvo_e)
        return
    calibrar_parametros_rapido(modo_completo=modo_completo_cli)
    print("▶️ Agora basta rodar: python main.py")


if __name__ == "__main__":
    main()


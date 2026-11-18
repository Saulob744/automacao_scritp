"""
main_background.py
==================
Automação PDTIC 100% em Segundo Plano (Background / Oculto).

Vantagens:
- NÃO mexe no mouse e NÃO aperta teclas físicas do teclado (sua tela fica 100% livre).
- Lê os dados diretamente do arquivo Excel baixado na raiz do projeto:
  'Campos Necessários para Sistema da Tecpar - Consolidado.xlsx'
- Usa a guia do Nex Privacy que JÁ ESTÁ ABERTA no seu navegador (sem recarregar e sem perder a sessão).
- Comunicação inteligente:
  * Canal A (CDP - Porta 9222): Conexão direta caso o Chrome esteja com depuração remota.
  * Canal B (Ponte - Porta 8765): Conexão local em tempo real via extensão de dados/extensao_background
    ou comando de 1 linha no Console [F12] da aba aberta.
"""

from __future__ import annotations

import argparse
import os
from pathlib import Path
import subprocess
import sys
import time

from src.automacao_background import (
    ARQUIVO_LOG_BG,
    DELAY_ADICIONAR_MS_PADRAO,
    NOME_ARQUIVO_EXCEL_PADRAO,
    PASTA_EXTENSAO_BG,
    PORTA_CDP_PADRAO,
    PORTA_PONTE_LOCAL,
    TAMANHO_LOTE_PADRAO,
    TIMEOUT_CONEXAO_PADRAO,
    URL_PADRAO_NEX_PRIVACY,
    executar_automacao_segundo_plano,
    localizar_arquivo_planilha_xlsx,
)


def iniciar_chrome_com_cdp(porta_cdp: int = PORTA_CDP_PADRAO, url: str = URL_PADRAO_NEX_PRIVACY) -> None:
    """Inicia uma instância do Google Chrome com depuração remota habilitada."""
    caminhos_chrome = [
        Path(os.environ.get("PROGRAMFILES", "C:\\Program Files")) / "Google" / "Chrome" / "Application" / "chrome.exe",
        Path(os.environ.get("PROGRAMFILES(X86)", "C:\\Program Files (x86)")) / "Google" / "Chrome" / "Application" / "chrome.exe",
        Path(os.environ.get("LOCALAPPDATA", "")) / "Google" / "Chrome" / "Application" / "chrome.exe",
    ]
    executavel = next((c for c in caminhos_chrome if c.is_file()), None)
    if not executavel:
        print("❌ Executável do Google Chrome não encontrado nos caminhos padrão.")
        return

    cmd = [
        str(executavel),
        f"--remote-debugging-port={porta_cdp}",
        url,
    ]
    print(f"🚀 Iniciando Chrome com porta CDP {porta_cdp}: {' '.join(cmd)}")
    subprocess.Popen(cmd)
    print("✅ Comando enviado! Aguarde o Chrome carregar a guia do sistema.")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Automação PDTIC 100% Oculta em 2º Plano (Usa a guia já aberta sem interferir na tela)."
    )
    parser.add_argument(
        "--bloco",
        "--executar",
        dest="bloco",
        type=str,
        default=None,
        help="Bloco ou página específica: 'todas', 'pagina3', '3.1', '3.2', 'pagina4', '4.1', '4.2', '4.3', 'pagina5', '5.1', '5.3', 'pagina6'.",
    )
    parser.add_argument(
        "--pagina3",
        "--diagnostico",
        dest="pagina3",
        action="store_true",
        help="Executa apenas a Página 3 ('3. Diagnóstico': 3.1 SWOT + 3.2 Inventário).",
    )
    parser.add_argument(
        "--pagina4",
        "--metas-acoes",
        dest="pagina4",
        action="store_true",
        help="Executa apenas a Página 4 ('4. Metas e Planos Obrigatórios': 4.1, 4.2 e 4.3).",
    )
    parser.add_argument(
        "--pagina5",
        "--capacitacao",
        dest="pagina5",
        action="store_true",
        help="Executa apenas a Página 5 ('5. Capacitação, Comunicação e Ações Específicas': 5.1 e 5.3).",
    )
    parser.add_argument(
        "--pagina6",
        "--monitoramento",
        dest="pagina6",
        action="store_true",
        help="Executa apenas a Página 6 ('6. Monitoramento, Anexos e Revisão': KPIs).",
    )
    parser.add_argument(
        "--tudo",
        "--todas",
        dest="tudo",
        action="store_true",
        help="Executa todas as etapas (Páginas 3, 4, 5 e 6) em sequência.",
    )
    parser.add_argument(
        "--limite",
        "--linhas",
        "-n",
        dest="limite",
        type=int,
        default=None,
        help="Limita a quantidade de linhas por bloco para teste rápido (ex: --limite 3).",
    )
    parser.add_argument(
        "--planilha",
        "--xlsx",
        dest="planilha",
        type=str,
        default=None,
        help=f"Caminho do arquivo Excel (padrão: '{NOME_ARQUIVO_EXCEL_PADRAO}').",
    )
    parser.add_argument(
        "--porta",
        dest="porta",
        type=int,
        default=PORTA_PONTE_LOCAL,
        help=f"Porta local da ponte oculta (padrão: {PORTA_PONTE_LOCAL}).",
    )
    parser.add_argument(
        "--porta-cdp",
        dest="porta_cdp",
        type=int,
        default=PORTA_CDP_PADRAO,
        help=f"Porta de depuração do Chrome CDP (padrão: {PORTA_CDP_PADRAO}).",
    )
    parser.add_argument(
        "--timeout",
        dest="timeout",
        type=int,
        default=TIMEOUT_CONEXAO_PADRAO,
        help=f"Tempo limite em segundos para conexão (padrão: {TIMEOUT_CONEXAO_PADRAO}s).",
    )
    parser.add_argument(
        "--iniciar-chrome",
        dest="iniciar_chrome",
        action="store_true",
        help="Abre o Google Chrome na porta CDP 9222 com o sistema aberto.",
    )
    parser.add_argument(
        "--analisar",
        "--diagnostico-dom",
        dest="analisar",
        action="store_true",
        help="Apenas analisa todos os campos, capacidades e estado atual do DOM vs Planilha sem alterar nada.",
    )
    parser.add_argument(
        "--delay-ms",
        dest="delay_ms",
        type=int,
        default=DELAY_ADICIONAR_MS_PADRAO,
        help=f"Tempo de espera (ms) após criar cada novo card no DOM (padrão: {DELAY_ADICIONAR_MS_PADRAO}ms).",
    )
    parser.add_argument(
        "--lote",
        dest="lote",
        type=int,
        default=TAMANHO_LOTE_PADRAO,
        help=f"Quantidade de linhas enviadas por lote (padrão: {TAMANHO_LOTE_PADRAO}).",
    )
    parser.add_argument(
        "--limpar",
        dest="limpar",
        action="store_true",
        help="Limpa linhas existentes de cada bloco antes de colar os dados da planilha.",
    )
    return parser.parse_args()


def menu_interativo() -> tuple[str, int | None]:
    """Exibe um menu intuitivo quando o script é executado sem parâmetros."""
    print("\n" + "=" * 78)
    print("🤖 AUTOMAÇÃO PDTIC — MODO SEGUNDO PLANO (BACKGROUND / 100% OCULTO)")
    print("=" * 78)
    print("Este modo lê o arquivo Excel baixado e preenche o sistema aberto no Chrome")
    print("SEM mexer no mouse, SEM apertar teclas e SEM roubar o foco da sua tela!\n")
    print("Escolha o que deseja executar:")
    print("  [1] Executar TUDO (Páginas 3, 4, 5 e 6 em sequência)")
    print("  [2] Página 3 (3.1 SWOT + 3.2 Inventário de necessidades)")
    print("  [3] Página 4 (4.1 Metas e Ações + 4.2 Plano Orçamentário + 4.3 Gestão de Riscos)")
    print("  [4] Página 5 (5.1 Capacitação + 5.3 Ações Específicas)")
    print("  [5] Página 6 (KPIs / Monitoramento)")
    print("  [6] Bloco Específico (ex: 4.1, 4.2, 4.3, 3.1, 3.2, 5.1, 5.3)")
    print("  [7] Abrir Google Chrome com porta CDP 9222 ativada")
    print("  [0] Sair")
    print("-" * 78)

    try:
        opcao = input("Digite a opção desejada [1-7]: ").strip()
    except (EOFError, KeyboardInterrupt):
        sys.exit(0)

    if opcao == "1":
        alvo = "todas"
    elif opcao == "2":
        alvo = "pagina3"
    elif opcao == "3":
        alvo = "pagina4"
    elif opcao == "4":
        alvo = "pagina5"
    elif opcao == "5":
        alvo = "pagina6"
    elif opcao == "6":
        try:
            bloco_dig = input("Digite o número do bloco (ex: 4.1, 4.2, 4.3, 3.1): ").strip()
            alvo = bloco_dig if bloco_dig else "todas"
        except (EOFError, KeyboardInterrupt):
            sys.exit(0)
    elif opcao == "7":
        iniciar_chrome_com_cdp()
        sys.exit(0)
    elif opcao == "0":
        sys.exit(0)
    else:
        print("Opção inválida. Executando 'todas'.")
        alvo = "todas"

    limite = None
    try:
        lim_str = input("Deseja limitar a quantidade de linhas para teste? (Pressione Enter para processar todas): ").strip()
        if lim_str.isdigit() and int(lim_str) > 0:
            limite = int(lim_str)
            print(f"⚠️ Limite ativado: apenas as primeiras {limite} linha(s) de cada bloco serão enviadas.")
    except (EOFError, KeyboardInterrupt):
        pass

    return alvo, limite


def main() -> None:
    args = parse_args()

    if args.iniciar_chrome:
        iniciar_chrome_com_cdp(porta_cdp=args.porta_cdp)
        return

    # Se nenhum argumento de execução foi passado, abre o menu interativo
    nenhum_alvo_passado = not any([
        args.bloco,
        args.pagina3,
        args.pagina4,
        args.pagina5,
        args.pagina6,
        args.tudo,
        args.analisar,
    ])

    if nenhum_alvo_passado and sys.stdin.isatty():
        alvo, limite = menu_interativo()
    else:
        if args.analisar and not any([args.bloco, args.pagina3, args.pagina4, args.pagina5, args.pagina6, args.tudo]):
            alvo = "analisar"
        elif args.bloco:
            alvo = args.bloco
        elif args.pagina3 and not args.tudo:
            alvo = "pagina3"
        elif args.pagina4 and not args.tudo:
            alvo = "pagina4"
        elif args.pagina5 and not args.tudo:
            alvo = "pagina5"
        elif args.pagina6 and not args.tudo:
            alvo = "pagina6"
        else:
            alvo = "todas"
        limite = args.limite

    # Localiza o arquivo Excel
    caminho_planilha = localizar_arquivo_planilha_xlsx(args.planilha)
    nome_exib_planilha = caminho_planilha.name if caminho_planilha else NOME_ARQUIVO_EXCEL_PADRAO

    print("\n" + "=" * 78)
    print("👻 AUTOMAÇÃO PDTIC 100% OCULTA (SEGUNDO PLANO / BACKGROUND)")
    print("=" * 78)
    print(f"🎯 Alvo da Execução:  {alvo.upper()}")
    print(f"📊 Limite de Linhas:  {limite if limite else 'TODAS AS LINHAS'}")
    print(f"📗 Arquivo Excel:     {nome_exib_planilha}")
    print(f"🌐 Guia do Sistema:   {URL_PADRAO_NEX_PRIVACY}")
    print(f"🧩 Extensão Oculta:   {PASTA_EXTENSAO_BG}")
    print(f"🪵 Log de Auditoria:  {ARQUIVO_LOG_BG}")
    print(f"⏱️  Espera / Lote:     {args.delay_ms}ms por card | Lote de {args.lote} linhas | Timeout={args.timeout}s")
    print("🖥️  Sua Tela:          100% LIVRE — Mouse e teclado não são afetados!")
    print("=" * 78 + "\n")

    sucesso = executar_automacao_segundo_plano(
        alvo=alvo,
        caminho_planilha=str(caminho_planilha) if caminho_planilha else None,
        limite=limite,
        porta_cdp=args.porta,
        timeout_conexao=args.timeout,
        delay_add_ms=args.delay_ms,
        tamanho_lote=args.lote,
        limpar_antes=args.limpar,
        apenas_analisar=args.analisar,
    )

    if sucesso:
        print("\n" + "🎉" * 20)
        print("✅ [SUCESSO] Preenchimento em segundo plano concluído com êxito!")
        print("🎉" * 20 + "\n")
    else:
        print("\n❌ A execução foi interrompida antes do final. Verifique os avisos acima.")


if __name__ == "__main__":
    main()

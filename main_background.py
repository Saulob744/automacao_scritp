"""
main_background.py
Versão 100% Oculta / 2º Plano (SEPARADA do `main.py` que ocupa a tela principal):
- NUNCA abre nenhuma janela na sua tela e NUNCA traz o navegador para a frente!
- Usa diretamente a guia `'https://parana.nexprivacy.com/admin/institutional-documents/100/edit'`
  que JÁ ESTÁ ABERTA no seu Google Chrome (mesmo minimizada, coberta ou em outro monitor).
- Lê a planilha `'Campos Necessários para Sistema da Tecpar - SESP'` de forma 100% invisível em memória.
"""

import argparse
from src.automacao_background import (
    ARQUIVO_LOG_DEBUG,
    NOME_PADRAO_PLANILHA_DRIVE,
    PASTA_EXTENSAO_BG,
    PORTA_PONTE_LOCAL,
    TIMEOUT_CONEXAO_PADRAO,
    URL_PADRAO_NEX_PRIVACY,
    URL_PADRAO_PLANILHA_DRIVE,
    executar_automacao_segundo_plano,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Automação PDTIC 100% Oculta em 2º Plano (Usa a guia já aberta do Nex Privacy sem subir na tela)."
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
        help="Executa apenas a Página 3 ('3. Diagnóstico': 3.1 SWOT + 3.2 Inventário) em modo oculto.",
    )
    parser.add_argument(
        "--pagina4",
        "--metas-acoes",
        dest="pagina4",
        action="store_true",
        help="Executa apenas a Página 4 ('4. Metas e Planos Obrigatórios': 4.1, 4.2 e 4.3) em modo oculto.",
    )
    parser.add_argument(
        "--pagina5",
        "--capacitacao",
        dest="pagina5",
        action="store_true",
        help="Executa apenas a Página 5 ('5. Capacitação, Comunicação e Ações Específicas': 5.1 e 5.3) em modo oculto.",
    )
    parser.add_argument(
        "--pagina6",
        "--monitoramento",
        dest="pagina6",
        action="store_true",
        help="Executa apenas a Página 6 ('6. Monitoramento, Anexos e Revisão': KPIs) em modo oculto.",
    )
    parser.add_argument(
        "--tudo",
        "--todas",
        dest="tudo",
        action="store_true",
        help="Executa todas as etapas (Páginas 3, 4, 5 e 6) em modo oculto.",
    )
    parser.add_argument(
        "--limite",
        "--linhas",
        "-n",
        dest="limite",
        type=int,
        default=None,
        help="Limita a quantidade de linhas por bloco para teste rápido (ex: --limite 5).",
    )
    parser.add_argument(
        "--planilha",
        "--xlsx",
        dest="planilha",
        type=str,
        default=None,
        help="Caminho opcional caso queira apontar um arquivo .xlsx local específico.",
    )
    parser.add_argument(
        "--porta",
        dest="porta",
        type=int,
        default=PORTA_PONTE_LOCAL,
        help="Porta local da ponte oculta (padrão: 8765).",
    )
    parser.add_argument(
        "--timeout",
        dest="timeout",
        type=int,
        default=TIMEOUT_CONEXAO_PADRAO,
        help=f"Tempo limite em segundos para aguardar resposta de cada etapa antes de parar e exibir diagnóstico (padrão: {TIMEOUT_CONEXAO_PADRAO}s).",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    print("=" * 78)
    print("👻 AUTOMAÇÃO PDTIC 100% OCULTA (USA A GUIA JÁ ABERTA SEM SUBIR NA TELA)")
    print("=" * 78)
    print(f"📄 Guia já aberta:   {URL_PADRAO_NEX_PRIVACY}")
    print(f"📗 Planilha Drive:   '{NOME_PADRAO_PLANILHA_DRIVE}'")
    print(f"🧩 Extensão Oculta:  {PASTA_EXTENSAO_BG}")
    print(f"🪵 Log de Debug:     {ARQUIVO_LOG_DEBUG}")
    print(f"⏱️  Timer / Timeout:  {args.timeout}s por etapa (para e avisa se travar)")
    print("🖥️  Tela / Monitores: Nenhuma janela será aberta ou trazida para a frente!")
    print("=" * 78)

    if args.bloco:
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

    executar_automacao_segundo_plano(
        alvo=alvo,
        caminho_planilha=args.planilha,
        limite=args.limite,
        porta_cdp=args.porta,
        timeout_conexao=args.timeout,
    )


if __name__ == "__main__":
    main()

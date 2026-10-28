"""
Utilitário Mapeador de Coordenadas (X, Y)
Execute este script para descobrir a posição exata de cada campo/botão na sua tela.
Para capturar e fixar um ponto no terminal, posicione o mouse e aperte a tecla 'F8'.
Para sair, aperte 'ESC' ou Ctrl+C no terminal.
"""

import time
from src.driver_windows import obter_posicao_mouse, tecla_pressionada


def iniciar_mapeador() -> None:
    print("=" * 65)
    print("🎯 MAPEADOR DE COORDENADAS DA TELA")
    print("=" * 65)
    print("- Mova o mouse até o campo ou botão desejado no seu sistema.")
    print("- Pressione [F8] para registrar/salvar a coordenada no histórico abaixo.")
    print("- Pressione [ESC] (ou Ctrl+C) para encerrar o mapeador.")
    print("=" * 65)

    pontos_salvos = []
    try:
        while True:
            if tecla_pressionada("esc"):
                print("\n\nEncerrando o mapeador...")
                break

            x, y = obter_posicao_mouse()
            print(f"\rPosição atual do mouse ->  X: {x:4d} | Y: {y:4d}   (Aperte F8 para gravar)", end="", flush=True)

            if tecla_pressionada("f8"):
                numero = len(pontos_salvos) + 1
                pontos_salvos.append((x, y))
                print(f"\n✅ Ponto #{numero} gravado: ({x}, {y})")
                time.sleep(0.4)  # Evita múltiplos registros no mesmo aperto

            time.sleep(0.05)
    except KeyboardInterrupt:
        print("\n\nMapeador interrompido pelo usuário.")

    if pontos_salvos:
        print("\n📋 Resumo dos pontos gravados (copie para o arquivo src/config.py):")
        for idx, (px, py) in enumerate(pontos_salvos, start=1):
            print(f'    "ponto_{idx}": ({px}, {py}),')


if __name__ == "__main__":
    iniciar_mapeador()
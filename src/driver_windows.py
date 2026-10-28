"""
Driver Híbrido de Automação para Windows.
Utiliza PyAutoGUI / Pyperclip / Keyboard caso estejam instalados no ambiente,
ou utiliza automaticamente a API nativa do Windows (ctypes.windll.user32 / kernel32)
para que o robô funcione imediatamente mesmo em redes com proxy corporativo restrito.
"""

import ctypes
from ctypes import wintypes
import os
import sys
import threading
import time
from typing import Tuple

try:
    import pyautogui  # type: ignore

    TEM_PYAUTOGUI = True
except ImportError:
    TEM_PYAUTOGUI = False

try:
    import pyperclip  # type: ignore

    TEM_PYPERCLIP = True
except ImportError:
    TEM_PYPERCLIP = False

try:
    import keyboard  # type: ignore

    TEM_KEYBOARD = True
except ImportError:
    TEM_KEYBOARD = False


# ============================================================================
# CONSTANTES E ESTRUTURAS DA API NATIVA DO WINDOWS (Win32 user32 / kernel32)
# ============================================================================
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
gdi32 = ctypes.windll.gdi32

try:
    user32.SetProcessDPIAware()
except Exception:
    pass

MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
MOUSEEVENTF_RIGHTDOWN = 0x0008
MOUSEEVENTF_RIGHTUP = 0x0010
MOUSEEVENTF_WHEEL = 0x0800
WHEEL_DELTA = 120

KEYEVENTF_KEYUP = 0x0002
CF_UNICODETEXT = 13
GMEM_MOVEABLE = 0x0002

user32.mouse_event.argtypes = [
    wintypes.DWORD,
    wintypes.DWORD,
    wintypes.DWORD,
    wintypes.DWORD,
    ctypes.c_size_t,
]
user32.mouse_event.restype = None

# Códigos de Teclas Virtuais (Virtual-Key Codes) do Windows
VK_MAP = {
    "backspace": 0x08,
    "tab": 0x09,
    "enter": 0x0D,
    "shift": 0x10,
    "ctrl": 0x11,
    "alt": 0x12,
    "pause": 0x13,
    "esc": 0x1B,
    "escape": 0x1B,
    "space": 0x20,
    "pageup": 0x21,
    "pagedown": 0x22,
    "end": 0x23,
    "home": 0x24,
    "left": 0x25,
    "up": 0x26,
    "right": 0x27,
    "down": 0x28,
    "delete": 0x2E,
    "f1": 0x70,
    "f2": 0x71,
    "f3": 0x72,
    "f4": 0x73,
    "f5": 0x74,
    "f6": 0x75,
    "f7": 0x76,
    "f8": 0x77,
    "f9": 0x78,
    "f10": 0x79,
    "f11": 0x7A,
    "f12": 0x7B,
}


class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]


class FailSafeException(Exception):
    """Disparada quando o mouse é movido para o canto superior esquerdo da tela (0, 0)."""


_MONITOR_PARADA_ATIVO = False
_PARADA_EMERGENCIA_ACIONADA = False
_ESC_SIMULADO_ATE = 0.0
_THREAD_MONITOR: threading.Thread | None = None


def _liberar_modificadores_teclado() -> None:
    """Garante que Ctrl, Shift e Alt sejam soltos ao abortar a execução."""
    for vk_mod in (0x10, 0x11, 0x12):
        try:
            user32.keybd_event(vk_mod, 0, KEYEVENTF_KEYUP, 0)
        except Exception:
            pass


def _acionar_parada_imediata(motivo: str) -> None:
    """Interrompe imediatamente toda a automação (mesmo no meio de sleep/scroll/clique)."""
    global _PARADA_EMERGENCIA_ACIONADA
    if _PARADA_EMERGENCIA_ACIONADA:
        os._exit(0)
    _PARADA_EMERGENCIA_ACIONADA = True
    _liberar_modificadores_teclado()
    sys.stdout.write(
        f"\n\n🛑 [PARADA DE EMERGÊNCIA IMEDIATA] {motivo}\n"
        f"   Algoritmo encerrado instantaneamente com segurança.\n"
    )
    sys.stdout.flush()
    os._exit(0)


def _checar_teclas_parada_emergencia() -> None:
    """
    Verifica em nível de hardware (Win32 GetAsyncKeyState) se o usuário apertou
    F12, ESC ou PAUSE para parar tudo imediatamente (sem precisar de Ctrl+C no terminal).
    """
    if not _MONITOR_PARADA_ATIVO:
        return
    if _PARADA_EMERGENCIA_ACIONADA:
        os._exit(0)

    # F12 (0x7B) e PAUSE (0x13) nunca são enviados pelo robô -> parada instantânea
    if bool(user32.GetAsyncKeyState(0x7B) & 0x8000):
        _acionar_parada_imediata("Tecla [F12] pressionada pelo usuário!")
    if bool(user32.GetAsyncKeyState(0x13) & 0x8000):
        _acionar_parada_imediata("Tecla [PAUSE] pressionada pelo usuário!")

    # ESC (0x1B) -> ignora apenas durante os ~120ms em que o próprio robô envia ESC para fechar Ctrl+F
    if time.perf_counter() > _ESC_SIMULADO_ATE:
        if bool(user32.GetAsyncKeyState(0x1B) & 0x8000):
            _acionar_parada_imediata("Tecla [ESC] pressionada pelo usuário!")


def _loop_watchdog_parada() -> None:
    """Thread em segundo plano que monitora a tecla de parada a cada 15ms."""
    while True:
        try:
            _checar_teclas_parada_emergencia()
        except Exception:
            pass
        time.sleep(0.015)


def iniciar_monitor_parada_emergencia() -> None:
    """
    Ativa o mecanismo de defesa em segundo plano: basta apertar [ESC] ou [F12]
    em qualquer janela que o algoritmo para tudo instantaneamente.
    """
    global _MONITOR_PARADA_ATIVO, _THREAD_MONITOR
    # Limpa estados anteriores das teclas antes de armar o monitor
    user32.GetAsyncKeyState(0x1B)
    user32.GetAsyncKeyState(0x7B)
    user32.GetAsyncKeyState(0x13)
    _MONITOR_PARADA_ATIVO = True
    if _THREAD_MONITOR is None or not _THREAD_MONITOR.is_alive():
        _THREAD_MONITOR = threading.Thread(target=_loop_watchdog_parada, daemon=True)
        _THREAD_MONITOR.start()


def _codigo_vk(tecla: str) -> int:
    t = tecla.lower().strip()
    if t in VK_MAP:
        return VK_MAP[t]
    if len(t) == 1:
        return ord(t.upper())
    raise ValueError(f"Tecla não reconhecida no mapeamento: '{tecla}'")


def obter_posicao_mouse() -> Tuple[int, int]:
    """Retorna a coordenada atual (X, Y) do ponteiro do mouse na tela."""
    if TEM_PYAUTOGUI:
        pos = pyautogui.position()
        return int(pos.x), int(pos.y)
    pt = POINT()
    user32.GetCursorPos(ctypes.byref(pt))
    return int(pt.x), int(pt.y)


def verificar_failsafe() -> None:
    """
    Verifica imediatamente:
    1. Tecla única de parada de emergência ([ESC], [F12] ou [PAUSE]).
    2. Canto superior esquerdo (0..3, 0..3) do monitor principal (Fail-Safe do mouse).
    """
    _checar_teclas_parada_emergencia()
    if TEM_PYAUTOGUI:
        pos = pyautogui.position()
        if 0 <= int(pos.x) <= 3 and 0 <= int(pos.y) <= 3:
            _acionar_parada_imediata("Mouse posicionado no canto superior esquerdo (0, 0)!")
        return

    pt = POINT()
    ok = user32.GetCursorPos(ctypes.byref(pt))
    if ok and 0 <= pt.x <= 3 and 0 <= pt.y <= 3:
        _acionar_parada_imediata("Mouse posicionado no canto superior esquerdo (0, 0)!")


def tecla_pressionada(tecla: str) -> bool:
    """Verifica em tempo real se uma tecla (ex: 'esc', 'f8') está sendo pressionada."""
    if TEM_KEYBOARD:
        try:
            return bool(keyboard.is_pressed(tecla))
        except Exception:
            pass
    vk = _codigo_vk(tecla)
    estado = user32.GetAsyncKeyState(vk)
    return bool(estado & 0x8000)


def aguardar_tecla_f8() -> Tuple[int, int]:
    """
    Aguarda o usuário posicionar o mouse e pressionar [F8], exibindo as coordenadas (X, Y) em tempo real.
    Retorna a tupla (x, y) capturada no instante do [F8].
    """
    while tecla_pressionada("f8"):
        time.sleep(0.03)

    while True:
        x, y = obter_posicao_mouse()
        sys.stdout.write(f"\r   📍 Mouse atual: X={x:4d}, Y={y:4d} (Aperte [F8] para confirmar)   ")
        sys.stdout.flush()

        if tecla_pressionada("esc") or tecla_pressionada("f12"):
            _acionar_parada_imediata("Captura [F8] cancelada pelo usuário via [ESC]/[F12]!")

        if tecla_pressionada("f8"):
            px, py = obter_posicao_mouse()
            while tecla_pressionada("f8"):
                time.sleep(0.03)
            sys.stdout.write("\n")
            sys.stdout.flush()
            time.sleep(0.15)
            return int(px), int(py)

        time.sleep(0.03)


def mover_e_clicar(x: int, y: int, cliques: int = 1, botao: str = "left") -> None:
    """Move o cursor até (X, Y) e realiza o clique."""
    verificar_failsafe()
    if TEM_PYAUTOGUI:
        pyautogui.click(x=x, y=y, clicks=cliques, button=botao)
        return

    user32.SetCursorPos(int(x), int(y))
    time.sleep(0.15)
    verificar_failsafe()

    down_flag = MOUSEEVENTF_LEFTDOWN if botao == "left" else MOUSEEVENTF_RIGHTDOWN
    up_flag = MOUSEEVENTF_LEFTUP if botao == "left" else MOUSEEVENTF_RIGHTUP

    for _ in range(cliques):
        verificar_failsafe()
        user32.mouse_event(down_flag, 0, 0, 0, 0)
        time.sleep(0.06)
        user32.mouse_event(up_flag, 0, 0, 0, 0)
        time.sleep(0.10)


def rolar_pagina(passos_scroll: int, x: int | None = None, y: int | None = None) -> None:
    """Rola a roda do mouse passo a passo (valores negativos rolam para baixo, positivos para cima)."""
    verificar_failsafe()
    if x is not None and y is not None:
        user32.SetCursorPos(int(x), int(y))
        time.sleep(0.06)
    if passos_scroll == 0:
        return
    sinal = 1 if passos_scroll > 0 else -1
    qtd = abs(int(passos_scroll))
    if TEM_PYAUTOGUI:
        for _ in range(qtd):
            verificar_failsafe()
            pyautogui.scroll(int(sinal * WHEEL_DELTA))
            time.sleep(0.025)
        return
    delta_unitario = ctypes.c_ulong(int(sinal * WHEEL_DELTA) & 0xFFFFFFFF).value
    for _ in range(qtd):
        verificar_failsafe()
        user32.mouse_event(MOUSEEVENTF_WHEEL, 0, 0, delta_unitario, 0)
        time.sleep(0.025)


def pressionar_tecla_unica(tecla: str, repeticoes: int = 1, intervalo: float = 0.08) -> None:
    """Pressiona e solta uma tecla específica."""
    global _ESC_SIMULADO_ATE
    verificar_failsafe()
    vk = _codigo_vk(tecla)
    eh_esc_interno = vk == 0x1B
    for _ in range(repeticoes):
        if eh_esc_interno:
            _ESC_SIMULADO_ATE = time.perf_counter() + 0.16
        else:
            verificar_failsafe()
        user32.keybd_event(vk, 0, 0, 0)
        time.sleep(0.03)
        user32.keybd_event(vk, 0, KEYEVENTF_KEYUP, 0)
        if eh_esc_interno:
            _ESC_SIMULADO_ATE = time.perf_counter() + 0.12
            user32.GetAsyncKeyState(0x1B)
        time.sleep(intervalo)


def pressionar_shift_tab(repeticoes: int = 1, intervalo: float = 0.08) -> None:
    """Executa Shift+Tab N vezes com tempo de retenção seguro para navegadores."""
    verificar_failsafe()
    vk_shift = _codigo_vk("shift")
    vk_tab = _codigo_vk("tab")
    for _ in range(repeticoes):
        verificar_failsafe()
        user32.keybd_event(vk_shift, 0, 0, 0)
        time.sleep(0.03)
        user32.keybd_event(vk_tab, 0, 0, 0)
        time.sleep(0.03)
        user32.keybd_event(vk_tab, 0, KEYEVENTF_KEYUP, 0)
        time.sleep(0.03)
        user32.keybd_event(vk_shift, 0, KEYEVENTF_KEYUP, 0)
        time.sleep(intervalo)


def atalho_teclado(*teclas: str) -> None:
    """Executa uma combinação de teclas (ex: atalho_teclado('ctrl', 'v'))."""
    verificar_failsafe()
    vks = [_codigo_vk(t) for t in teclas]
    for vk in vks:
        user32.keybd_event(vk, 0, 0, 0)
        time.sleep(0.03)
    for vk in reversed(vks):
        user32.keybd_event(vk, 0, KEYEVENTF_KEYUP, 0)
        time.sleep(0.03)



# Configuração dos tipos de ponteiro do Clipboard na API Win32 (64-bit safe)
kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
kernel32.GlobalLock.restype = wintypes.LPVOID
kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
kernel32.GlobalUnlock.restype = wintypes.BOOL
user32.OpenClipboard.argtypes = [wintypes.HWND]
user32.OpenClipboard.restype = wintypes.BOOL
user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
user32.SetClipboardData.restype = wintypes.HANDLE
user32.GetClipboardData.argtypes = [wintypes.UINT]
user32.GetClipboardData.restype = wintypes.HANDLE


def copiar_para_clipboard(texto: str) -> None:
    """Coloca um texto na área de transferência do Windows preservando acentuação."""
    if TEM_PYPERCLIP:
        pyperclip.copy(texto)
        return

    dados_bytes = (texto + "\0").encode("utf-16-le")
    if not user32.OpenClipboard(None):
        time.sleep(0.05)
        user32.OpenClipboard(None)
    try:
        user32.EmptyClipboard()
        h_mem = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(dados_bytes))
        ptr = kernel32.GlobalLock(h_mem)
        ctypes.memmove(ptr, dados_bytes, len(dados_bytes))
        kernel32.GlobalUnlock(h_mem)
        user32.SetClipboardData(CF_UNICODETEXT, h_mem)
    finally:
        user32.CloseClipboard()


def ler_do_clipboard() -> str:
    """Lê o texto atual da área de transferência do Windows."""
    if TEM_PYPERCLIP:
        return str(pyperclip.paste())

    if not user32.OpenClipboard(None):
        time.sleep(0.05)
        user32.OpenClipboard(None)
    try:
        h_mem = user32.GetClipboardData(CF_UNICODETEXT)
        if not h_mem:
            return ""
        ptr = kernel32.GlobalLock(h_mem)
        if not ptr:
            return ""
        try:
            return ctypes.wstring_at(ptr)
        finally:
            kernel32.GlobalUnlock(h_mem)
    finally:
        user32.CloseClipboard()

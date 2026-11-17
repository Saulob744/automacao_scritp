"""
Modo de Automação 100% Oculto / Segundo Plano (SEPARADO do `main.py`).

Inclui:
1. Sistema completo de `[DEBUG]` passo a passo (no terminal e em `dados/debug_background.log`).
2. Timers (Watchdog / Timeout) em TODAS as etapas: nunca fica travado infinitamente;
   se qualquer etapa não responder dentro do tempo limite, PARA IMEDIATAMENTE e exibe
   onde parou, qual erro ocorreu e como resolver.
3. Suporte simultâneo a:
   - Canal A (CDP Porta 9222): conecta direto na guia já aberta se o Chrome estiver com `--remote-debugging-port=9222`.
   - Canal B (Ponte / Extensão Porta 8765): conecta na guia já aberta via extensão `dados/extensao_background` ou Console (`F12`).
"""

from __future__ import annotations

import base64
from http.server import BaseHTTPRequestHandler, HTTPServer
import json
import os
from pathlib import Path
import re
import shutil
import socket
import sqlite3
import subprocess
import sys
import tempfile
import threading
import time
from typing import Any, Dict, List, Optional, Tuple
import urllib.parse
import urllib.request

import openpyxl

from src.automacao import (
    COLUNAS_ACOES_ESPECIFICAS_5_3 as COLUNAS_ACOES_ESPECIFICAS,
    COLUNAS_CAPACITACAO_5_1 as COLUNAS_CAPACITACAO,
    COLUNAS_GESTAO_RISCOS,
    COLUNAS_INVENTARIO_NECESSIDADES,
    COLUNAS_METAS_ACOES,
    COLUNAS_MONITORAMENTO_KPIS_6 as COLUNAS_MONITORAMENTO_KPIS,
    COLUNAS_PADRAO_SWOT,
    COLUNAS_PLANO_ORCAMENTARIO,
    buscar_valor_da_coluna,
    obter_dados_etapa_4_da_memoria,
    salvar_dados_etapa_4_em_memoria,
)
from src.config import (
    DADOS_DIR,
    carregar_mapeamento,
)
from src.leitor_dados import (
    extrair_tabela_de_texto_tsv,
    remover_acentos,
)

PORTA_PONTE_LOCAL = 8765
PORTA_CDP_PADRAO = 9222
TIMEOUT_CONEXAO_PADRAO = 25
TIMEOUT_COMANDO_PADRAO = 15
PASTA_EXTENSAO_BG = DADOS_DIR / "extensao_background"
ARQUIVO_LOG_BG = DADOS_DIR / "debug_background.log"
URL_PADRAO_NEX_PRIVACY = "https://parana.nexprivacy.com/admin/institutional-documents/100/edit"
URL_PADRAO_PLANILHA_DRIVE = "https://docs.google.com/spreadsheets/d/1lqrbuQwTY93KHhUP1WmyQR44oGhiD4gsXeVbASnZycs/edit"
ID_PADRAO_PLANILHA_DRIVE = "1lqrbuQwTY93KHhUP1WmyQR44oGhiD4gsXeVbASnZycs"
NOME_PADRAO_PLANILHA_DRIVE = "Campos Necessários para Sistema da Tecpar - Consolidado"
NOME_ARQUIVO_EXCEL_PADRAO = "Campos Necessários para Sistema da Tecpar - Consolidado.xlsx"


# =============================================================================
# 0. LOGGER DE DEBUG PASSO A PASSO + DIAGNÓSTICO DE PARADA
# =============================================================================


def log_debug(etapa: str, mensagem: str, nivel: str = "DEBUG") -> None:
    """Registra passo a passo no console e em `dados/debug_background.log` com timestamp."""
    ts = time.strftime("%H:%M:%S")
    icone = {
        "DEBUG": "🔎",
        "INFO": "ℹ️ ",
        "OK": "✅",
        "WARN": "⚠️ ",
        "ERRO": "❌",
        "TIMER": "⏱️ ",
    }.get(nivel.upper(), "🔎")
    linha = f"[{ts}] {icone} [{nivel.upper()} | {etapa}] {mensagem}"
    print(linha, flush=True)
    try:
        DADOS_DIR.mkdir(parents=True, exist_ok=True)
        with open(ARQUIVO_LOG_BG, "a", encoding="utf-8") as f:
            f.write(linha + "\n")
    except Exception:
        pass


def abortar_com_diagnostico(
    onde_parou: str,
    motivo_erro: str,
    detalhes_tecnicos: List[str],
    como_resolver: List[str],
) -> bool:
    """
    Interrompe imediatamente a execução quando um timer expira ou ocorre um erro,
    exibindo um relatório claro de ONDE parou, POR QUE parou e COMO resolver.
    """
    print("\n" + "!" * 78, flush=True)
    log_debug(onde_parou, f"PARADA AUTOMÁTICA DE SEGURANÇA: {motivo_erro}", nivel="ERRO")
    print("-" * 78)
    print(f"📍 ONDE PAROU:    {onde_parou}")
    print(f"❌ MOTIVO/ERRO:   {motivo_erro}")
    if detalhes_tecnicos:
        print("🧪 DIAGNÓSTICO TÉCNICO DO DEBUG:")
        for d in detalhes_tecnicos:
            print(f"   • {d}")
    if como_resolver:
        print("🛠️  COMO RESOLVER AGORA:")
        for idx, c in enumerate(como_resolver, start=1):
            print(f"   {idx}. {c}")
    print(f"📄 Log completo salvo em: {ARQUIVO_LOG_BG}")
    print("!" * 78 + "\n", flush=True)
    return False


def descobrir_urls_no_historico_chrome(
    termo_planilha: str = "Campos Necess",
) -> Dict[str, Any]:
    """
    Inspeciona de forma somente-leitura (via cópia temporária) o banco SQLite `History`
    dos perfis do Google Chrome do usuário para localizar a URL exata da planilha e do sistema.
    """
    base_chrome = Path(os.environ.get("LOCALAPPDATA", "")) / "Google" / "Chrome" / "User Data"
    resultado: Dict[str, Any] = {
        "perfil": "Profile 1",
        "url_planilha": URL_PADRAO_PLANILHA_DRIVE,
        "sheet_id": ID_PADRAO_PLANILHA_DRIVE,
        "titulo_planilha": NOME_PADRAO_PLANILHA_DRIVE,
        "url_sistema": URL_PADRAO_NEX_PRIVACY,
    }
    if not base_chrome.is_dir():
        return resultado

    termo_norm = remover_acentos(termo_planilha)
    melhor_visita_planilha = -1

    for pasta_perfil in sorted(base_chrome.iterdir()):
        if not pasta_perfil.is_dir():
            continue
        arq_hist = pasta_perfil / "History"
        if not arq_hist.is_file():
            continue

        tmp_db = Path(tempfile.gettempdir()) / f"agy_chrome_hist_{pasta_perfil.name}.db"
        try:
            shutil.copy2(arq_hist, tmp_db)
            con = sqlite3.connect(str(tmp_db))
            cur = con.cursor()
            rows = cur.execute(
                "SELECT url, title, last_visit_time FROM urls "
                "WHERE url LIKE '%docs.google.com/spreadsheets/d/%' "
                "OR url LIKE '%nexprivacy.com/admin/institutional-documents/%' "
                "ORDER BY last_visit_time DESC LIMIT 60"
            ).fetchall()
            con.close()
        except Exception:
            continue

        for url, title, visit_time in rows:
            url_str = str(url or "")
            title_str = str(title or "")
            v_time = int(visit_time or 0)

            if "nexprivacy.com/admin/institutional-documents/" in url_str and "/edit" in url_str:
                resultado["url_sistema"] = url_str

            if "docs.google.com/spreadsheets/d/" in url_str:
                m_id = re.search(r"docs\.google\.com/spreadsheets/d/([a-zA-Z0-9_-]+)", url_str)
                if not m_id:
                    continue
                sheet_id = m_id.group(1)
                url_limpa = f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"
                titulo_norm = remover_acentos(title_str)
                eh_alvo = (
                    termo_norm in titulo_norm
                    or ("TECPAR" in titulo_norm and "SESP" in titulo_norm)
                    or ("CAMPOS NECESSARIOS" in titulo_norm)
                )
                if eh_alvo and v_time > melhor_visita_planilha:
                    melhor_visita_planilha = v_time
                    resultado["url_planilha"] = url_limpa
                    resultado["sheet_id"] = sheet_id
                    resultado["titulo_planilha"] = title_str
                    resultado["perfil"] = pasta_perfil.name

    return resultado


# =============================================================================
# 1. JAVASCRIPT DE PREENCHIMENTO DIRETO NO DOM (USADO TANTO VIA CDP QUANTO VIA PONTE)
# =============================================================================

JS_HELPER_PREENCHER_REPEATER = r"""
(function(config) {
    const stepIndex = config.stepIndex;
    const repeaterName = config.repeaterName;
    const mode = config.mode;
    const items = config.items || [];

    if (stepIndex !== null && stepIndex !== undefined) {
        const navBtn = document.querySelector(`.wizard-nav button[data-step-index="${stepIndex}"]`);
        if (navBtn && !navBtn.classList.contains('active')) {
            navBtn.click();
        }
    }

    const container = document.querySelector(`[data-repeater="${repeaterName}"]`);
    const addBtn = document.querySelector(`button[data-repeater-add="${repeaterName}"]`);
    if (!container) {
        return {
            ok: false,
            erro: `Container HTML [data-repeater="${repeaterName}"] não foi encontrado na guia aberta (${location.href}).`
        };
    }

    function dispararEventos(el) {
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.dispatchEvent(new Event('blur', { bubbles: true }));
    }

    function normalizarTexto(s) {
        return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();
    }

    function definirValorElemento(el, valor) {
        if (!el) return false;
        const valStr = String(valor ?? '').trim();
        if (!valStr) return false;

        if (el.tagName.toLowerCase() === 'select') {
            const alvoNorm = normalizarTexto(valStr);
            let escolheu = false;
            for (const opt of Array.from(el.options)) {
                const optValNorm = normalizarTexto(opt.value);
                const optTxtNorm = normalizarTexto(opt.textContent);
                if (optValNorm === alvoNorm || optTxtNorm === alvoNorm || optTxtNorm.includes(alvoNorm) || alvoNorm.includes(optTxtNorm)) {
                    el.value = opt.value;
                    escolheu = true;
                    break;
                }
            }
            if (!escolheu) el.value = valStr;
            dispararEventos(el);
            return true;
        } else {
            const eraReadonly = el.hasAttribute('readonly');
            if (eraReadonly) el.removeAttribute('readonly');
            el.value = valStr;
            dispararEventos(el);
            if (eraReadonly) el.setAttribute('readonly', '');
            return true;
        }
    }

    function obterLinhasRepeater() {
        const rows = Array.from(container.querySelectorAll('[data-repeater-row], [data-repeater-item], .wizard-table-row, .repeater-item'));
        return rows.length > 0 ? rows : Array.from(container.children);
    }

    let preenchidos = 0;
    let detalhes = [];

    if (mode === 'single_column') {
        for (let i = 0; i < items.length; i++) {
            const texto = String(items[i] ?? '').trim();
            if (!texto) continue;
            let rows = obterLinhasRepeater();
            let alvoInput = null;
            let idxUsado = -1;
            for (let rIdx = 0; rIdx < rows.length; rIdx++) {
                const inp = rows[rIdx].querySelector('input.form-control, textarea.form-control');
                if (inp && String(inp.value || '').trim() === '') {
                    alvoInput = inp;
                    idxUsado = rIdx + 1;
                    break;
                }
            }
            if (!alvoInput) {
                if (!addBtn) {
                    return { ok: false, erro: `Botão [data-repeater-add="${repeaterName}"] não encontrado.` };
                }
                addBtn.click();
                rows = obterLinhasRepeater();
                const ultimaRow = rows[rows.length - 1];
                idxUsado = rows.length;
                alvoInput = ultimaRow ? ultimaRow.querySelector('input.form-control, textarea.form-control') : null;
            }
            if (alvoInput) {
                definirValorElemento(alvoInput, texto);
                preenchidos++;
                detalhes.push(`caixa #${idxUsado} (${alvoInput.name || 'input'}) <- "${texto.slice(0, 35)}"`);
            }
        }
        return { ok: true, preenchidos: preenchidos, totalLinhasDOM: obterLinhasRepeater().length, detalhes: detalhes };
    }

    for (let i = 0; i < items.length; i++) {
        const mapaCampos = items[i] || {};
        let rows = obterLinhasRepeater();
        let rowVazia = null;
        let idxUsado = -1;
        for (let rIdx = 0; rIdx < rows.length; rIdx++) {
            const r = rows[rIdx];
            const editaveis = Array.from(r.querySelectorAll('input.form-control:not([readonly]), textarea.form-control:not([readonly])'));
            const temTexto = editaveis.some(el => String(el.value || '').trim() !== '');
            if (!temTexto && editaveis.length > 0) {
                rowVazia = r;
                idxUsado = rIdx + 1;
                break;
            }
        }
        if (!rowVazia) {
            if (!addBtn) {
                return { ok: false, erro: `Botão [data-repeater-add="${repeaterName}"] não encontrado.` };
            }
            addBtn.click();
            rows = obterLinhasRepeater();
            rowVazia = rows[rows.length - 1];
            idxUsado = rows.length;
        }
        if (rowVazia) {
            let camposOk = 0;
            for (const [seletorRelativo, valorCampo] of Object.entries(mapaCampos)) {
                const valStr = String(valorCampo ?? '').trim();
                if (!valStr) continue;
                const elCampo = rowVazia.querySelector(seletorRelativo);
                if (elCampo && definirValorElemento(elCampo, valStr)) {
                    camposOk++;
                }
            }
            preenchidos++;
            detalhes.push(`sub-card #${idxUsado}: ${camposOk} campo(s) preenchido(s)`);
        }
    }
    return { ok: true, preenchidos: preenchidos, totalLinhasDOM: obterLinhasRepeater().length, detalhes: detalhes };
})
"""

JS_BRIDGE_CONSOLE = (
    r"""
(function() {
    if (window.__PDTIC_BRIDGE_RUNNING__) {
        console.log("[PDTIC 2º Plano] Ponte já está ativa nesta guia!");
        return;
    }
    window.__PDTIC_BRIDGE_RUNNING__ = true;
    const SERVIDOR = "http://127.0.0.1:8765";
    const fnPreencher = """
    + JS_HELPER_PREENCHER_REPEATER.strip()
    + r""";

    async function enviarResultado(id, resultado) {
        try {
            await fetch(`${SERVIDOR}/api/result`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ id, resultado })
            });
        } catch (e) {}
    }

    async function poll() {
        try {
            const r = await fetch(`${SERVIDOR}/api/poll`, { cache: 'no-store' });
            if (!r.ok) return;
            const tarefa = await r.json();
            if (!tarefa || !tarefa.id) return;

            if (tarefa.acao === 'ping') {
                const btn = document.querySelector('.wizard-nav button.active');
                await enviarResultado(tarefa.id, {
                    ok: true,
                    url: location.href,
                    title: document.title,
                    etapaAtiva: btn ? btn.textContent.trim() : '1. Identificação',
                    modo: 'ponte_direta_guia_aberta'
                });
            } else if (tarefa.acao === 'fill_repeater') {
                const res = fnPreencher(tarefa.payload);
                await enviarResultado(tarefa.id, res);
            }
        } catch (e) {}
    }

    setInterval(poll, 450);
    poll();
    console.log("✅ [PDTIC 2º Plano] Guia conectada! Você já pode minimizar ou cobrir esta janela.");
})();
"""
)


# =============================================================================
# 2. CANAL A: CLIENTE CDP DIRETO (PORTA 9222, SE DISPONÍVEL)
# =============================================================================


class ClienteCDP:
    """Cliente WebSocket leve (stdlib pura) com timeout estrito para cada chamada."""

    def __init__(self, ws_url: str, timeout: float = 10.0) -> None:
        parsed = urllib.parse.urlparse(ws_url)
        host = parsed.hostname or "127.0.0.1"
        port = parsed.port or 9222
        path = parsed.path or "/"
        if parsed.query:
            path += "?" + parsed.query

        self.sock = socket.create_connection((host, port), timeout=timeout)
        self.sock.settimeout(timeout)
        key = base64.b64encode(os.urandom(16)).decode("ascii")
        handshake = (
            f"GET {path} HTTP/1.1\r\n"
            f"Host: {host}:{port}\r\n"
            "Upgrade: websocket\r\n"
            "Connection: Upgrade\r\n"
            f"Sec-WebSocket-Key: {key}\r\n"
            "Sec-WebSocket-Version: 13\r\n\r\n"
        )
        self.sock.sendall(handshake.encode("ascii"))

        resposta = b""
        while b"\r\n\r\n" not in resposta:
            chunk = self.sock.recv(4096)
            if not chunk:
                raise RuntimeError("Conexão encerrada durante handshake WebSocket CDP.")
            resposta += chunk

        cabecalho = resposta.split(b"\r\n\r\n", 1)[0].decode("latin1", errors="ignore")
        if "101" not in cabecalho.splitlines()[0]:
            raise RuntimeError(f"Falha no handshake WebSocket CDP: {cabecalho.splitlines()[0]}")

        self._buffer_extra = resposta.split(b"\r\n\r\n", 1)[1]
        self._msg_id = 0

    def _recv_exact(self, n: int) -> bytes:
        dados = bytearray()
        if self._buffer_extra:
            take = min(n, len(self._buffer_extra))
            dados.extend(self._buffer_extra[:take])
            self._buffer_extra = self._buffer_extra[take:]
        while len(dados) < n:
            chunk = self.sock.recv(max(4096, n - len(dados)))
            if not chunk:
                raise RuntimeError("Conexão WebSocket CDP fechada pelo navegador.")
            dados.extend(chunk)
        return bytes(dados)

    def _ler_frame(self) -> str:
        while True:
            b1, b2 = self._recv_exact(2)
            opcode = b1 & 0x0F
            masked = (b2 & 0x80) != 0
            length = b2 & 0x7F
            if length == 126:
                length = int.from_bytes(self._recv_exact(2), "big")
            elif length == 127:
                length = int.from_bytes(self._recv_exact(8), "big")
            mask_key = self._recv_exact(4) if masked else b""
            payload = bytearray(self._recv_exact(length))
            if masked:
                for i in range(length):
                    payload[i] ^= mask_key[i % 4]
            if opcode == 0x8:
                raise RuntimeError("Aba do Chrome encerrou a conexão WebSocket.")
            if opcode == 0x9:
                self._enviar_frame(bytes(payload), opcode=0xA)
                continue
            if opcode in (0x1, 0x2):
                return payload.decode("utf-8", errors="replace")

    def _enviar_frame(self, dados: bytes, opcode: int = 0x1) -> None:
        header = bytearray([0x80 | (opcode & 0x0F)])
        length = len(dados)
        mask_key = os.urandom(4)
        if length < 126:
            header.append(0x80 | length)
        elif length < 65536:
            header.append(0x80 | 126)
            header.extend(length.to_bytes(2, "big"))
        else:
            header.append(0x80 | 127)
            header.extend(length.to_bytes(8, "big"))
        header.extend(mask_key)
        masked_payload = bytearray(dados)
        for i in range(length):
            masked_payload[i] ^= mask_key[i % 4]
        self.sock.sendall(header + masked_payload)

    def chamar(self, metodo: str, params: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        self._msg_id += 1
        req_id = self._msg_id
        pacote = json.dumps({"id": req_id, "method": metodo, "params": params or {}})
        self._enviar_frame(pacote.encode("utf-8"))
        while True:
            texto = self._ler_frame()
            msg = json.loads(texto)
            if msg.get("id") == req_id:
                if "error" in msg:
                    raise RuntimeError(f"Erro CDP ({metodo}): {msg['error']}")
                return msg.get("result") or {}

    def executar_js(self, expressao_js: str) -> Any:
        res = self.chamar(
            "Runtime.evaluate",
            {
                "expression": expressao_js,
                "returnByValue": True,
                "awaitPromise": True,
            },
        )
        if "exceptionDetails" in res:
            detalhe = res["exceptionDetails"]
            raise RuntimeError(f"Erro JS na aba: {detalhe}")
        return res.get("result", {}).get("value")

    def fechar(self) -> None:
        try:
            self.sock.close()
        except Exception:
            pass


def _tentar_conectar_via_cdp_9222(porta: int = PORTA_CDP_PADRAO) -> Tuple[Optional[ClienteCDP], str]:
    """
    Tenta conectar silenciosamente à porta CDP `9222` em até 1.5s (sem abrir nenhuma janela!).
    Retorna `(cliente_cdp, diagnostico)`.
    """
    url = f"http://127.0.0.1:{porta}/json"
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=1.5) as resp:
            dados = json.loads(resp.read().decode("utf-8"))
    except Exception as exc:
        return None, f"Porta CDP {porta} indisponível ({type(exc).__name__}: {exc})"

    abas = [d for d in dados if isinstance(d, dict) and d.get("type") == "page"]
    if not abas:
        return None, f"Porta CDP {porta} aberta, mas nenhuma aba 'page' foi encontrada."

    for aba in abas:
        u = str(aba.get("url") or "").lower()
        t = str(aba.get("title") or "").lower()
        if ("nexprivacy" in u or "institutional-documents" in u or "nex privacy" in t) and aba.get("webSocketDebuggerUrl"):
            try:
                cli = ClienteCDP(str(aba["webSocketDebuggerUrl"]), timeout=8.0)
                tem_form = bool(
                    cli.executar_js(
                        "Boolean(document.querySelector('#pdtic-form') || document.querySelector('.wizard-step'))"
                    )
                )
                if tem_form:
                    return cli, f"Conectado via CDP (porta {porta}) à guia '{aba.get('title')}' ({aba.get('url')})"
                cli.fechar()
            except Exception as exc:
                return None, f"Guia encontrada na porta {porta}, mas falhou ao abrir WebSocket ({exc})"

    resumo_abas = [f"{a.get('title')} ({a.get('url')})" for a in abas[:4]]
    return None, f"Porta {porta} ativa, porém a guia do Nex Privacy não está entre as abas: {resumo_abas}"


# =============================================================================
# 3. CANAL B: SERVIDOR PONTE OCULTA (PORTA 8765) COM LOG HTTP E TIMERS ESTRITOS
# =============================================================================


class ControladorExecucaoOculta:
    """
    Unifica o envio de comandos para a guia já aberta do Nex Privacy (seja via CDP 9222
    ou via Ponte/Extensão 8765), com logs de `[DEBUG]` e timers estritos em cada operação.
    """

    def __init__(self, porta_ponte: int = PORTA_PONTE_LOCAL, porta_cdp: int = PORTA_CDP_PADRAO) -> None:
        self.porta_ponte = porta_ponte
        self.porta_cdp = porta_cdp
        self.cliente_cdp: Optional[ClienteCDP] = None
        self.modo_ativo: str = "nenhum"

        self._lock = threading.Lock()
        self._tarefa_atual: Optional[Dict[str, Any]] = None
        self._resultados: Dict[str, Any] = {}
        self._evento_resultado = threading.Event()
        self._contador = 0
        self.qtd_polls_recebidos = 0
        self.qtd_bridge_baixados = 0
        self.ultimo_ip_cliente: Optional[str] = None
        self._httpd: Optional[HTTPServer] = None
        self._thread: Optional[threading.Thread] = None

    def iniciar_servidor_ponte(self) -> Tuple[bool, str]:
        pai = self

        class Handler(BaseHTTPRequestHandler):
            def log_message(self, format: str, *args: Any) -> None:
                pass

            def _enviar_cors(self) -> None:
                self.send_header("Access-Control-Allow-Origin", "*")
                self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
                self.send_header("Access-Control-Allow-Headers", "Content-Type")
                self.send_header("Access-Control-Allow-Private-Network", "true")
                self.send_header("Cache-Control", "no-store")

            def do_OPTIONS(self) -> None:
                pai.ultimo_ip_cliente = self.client_address[0]
                log_debug("HTTP-PONTE", f"Preflight OPTIONS recebido de {self.path}", nivel="DEBUG")
                self.send_response(204)
                self._enviar_cors()
                self.end_headers()

            def do_GET(self) -> None:
                pai.ultimo_ip_cliente = self.client_address[0]
                if self.path.startswith("/bridge.js"):
                    pai.qtd_bridge_baixados += 1
                    log_debug(
                        "HTTP-PONTE",
                        f"Script '/bridge.js' baixado pela guia do navegador! (total={pai.qtd_bridge_baixados})",
                        nivel="OK",
                    )
                    corpo = JS_BRIDGE_CONSOLE.encode("utf-8")
                    self.send_response(200)
                    self._enviar_cors()
                    self.send_header("Content-Type", "application/javascript; charset=utf-8")
                    self.end_headers()
                    self.wfile.write(corpo)
                    return

                if self.path.startswith("/api/poll"):
                    with pai._lock:
                        pai.qtd_polls_recebidos += 1
                        num_p = pai.qtd_polls_recebidos
                        tarefa = pai._tarefa_atual
                        pai._tarefa_atual = None
                    if num_p == 1 or (tarefa is not None):
                        log_debug(
                            "HTTP-PONTE",
                            f"Sinal '/api/poll' #{num_p} recebido da guia! "
                            f"Enviando comando: {tarefa.get('acao') if tarefa else 'nenhum (idle)'}",
                            nivel="DEBUG",
                        )
                    corpo = json.dumps(tarefa or {}).encode("utf-8")
                    self.send_response(200)
                    self._enviar_cors()
                    self.send_header("Content-Type", "application/json; charset=utf-8")
                    self.end_headers()
                    self.wfile.write(corpo)
                    return

                self.send_response(404)
                self._enviar_cors()
                self.end_headers()

            def do_POST(self) -> None:
                pai.ultimo_ip_cliente = self.client_address[0]
                if self.path.startswith("/api/result"):
                    tam = int(self.headers.get("Content-Length", "0"))
                    raw = self.rfile.read(tam) if tam > 0 else b"{}"
                    try:
                        dados = json.loads(raw.decode("utf-8"))
                        tid = str(dados.get("id") or "")
                        res_obj = dados.get("resultado")
                        log_debug(
                            "HTTP-PONTE",
                            f"Resposta recebida da guia para comando '{tid}': ok={res_obj.get('ok') if isinstance(res_obj, dict) else res_obj}",
                            nivel="DEBUG",
                        )
                        if tid:
                            with pai._lock:
                                pai._resultados[tid] = res_obj
                            pai._evento_resultado.set()
                    except Exception as exc:
                        log_debug("HTTP-PONTE", f"Erro ao decodificar resposta JSON da guia: {exc}", nivel="ERRO")
                    self.send_response(200)
                    self._enviar_cors()
                    self.send_header("Content-Type", "application/json")
                    self.end_headers()
                    self.wfile.write(b'{"ok":true}')
                    return

                self.send_response(404)
                self._enviar_cors()
                self.end_headers()

        try:
            self._httpd = HTTPServer(("127.0.0.1", self.porta_ponte), Handler)
            self._thread = threading.Thread(target=self._httpd.serve_forever, daemon=True)
            self._thread.start()
            return True, f"Servidor local ouvindo em http://127.0.0.1:{self.porta_ponte}"
        except Exception as exc:
            return False, f"Falha ao abrir porta local {self.porta_ponte}: {exc}"

    def fechar(self) -> None:
        if self.cliente_cdp is not None:
            self.cliente_cdp.fechar()
            self.cliente_cdp = None
        if self._httpd is not None:
            try:
                self._httpd.shutdown()
                self._httpd.server_close()
            except Exception:
                pass

    def enviar_comando_ponte(
        self,
        acao: str,
        payload: Optional[Dict[str, Any]] = None,
        timeout: float = float(TIMEOUT_COMANDO_PADRAO),
    ) -> Optional[Dict[str, Any]]:
        with self._lock:
            self._contador += 1
            tid = f"cmd_{int(time.time() * 1000)}_{self._contador}"
            self._evento_resultado.clear()
            pacote: Dict[str, Any] = {"id": tid, "acao": acao}
            if payload:
                pacote.update(payload)
            self._tarefa_atual = pacote

        fim = time.time() + timeout
        while time.time() < fim:
            if self._evento_resultado.wait(timeout=0.25):
                with self._lock:
                    if tid in self._resultados:
                        return self._resultados.pop(tid)
                self._evento_resultado.clear()
        return None

    def preencher_repeater(
        self,
        payload: Dict[str, Any],
        nome_exibicao: str,
        timeout: float = float(TIMEOUT_COMANDO_PADRAO),
    ) -> Dict[str, Any]:
        t0 = time.time()
        log_debug(
            nome_exibicao,
            f"Enviando preenchimento para a guia (repeater='{payload.get('repeaterName')}', "
            f"stepIndex={payload.get('stepIndex')}, itens={len(payload.get('items') or [])}, "
            f"modo={self.modo_ativo}, timeout={timeout:.0f}s)...",
            nivel="DEBUG",
        )

        if self.modo_ativo == "cdp" and self.cliente_cdp is not None:
            try:
                expressao = f"{JS_HELPER_PREENCHER_REPEATER}({json.dumps(payload, ensure_ascii=False)})"
                res = self.cliente_cdp.executar_js(expressao)
                dt = time.time() - t0
                if isinstance(res, dict):
                    res["_tempo_s"] = round(dt, 2)
                    return res
                return {"ok": False, "erro": f"Retorno inesperado do DOM via CDP: {res}"}
            except Exception as exc:
                return {"ok": False, "erro": f"Exceção CDP durante preenchimento de '{nome_exibicao}': {exc}"}

        res_ponte = self.enviar_comando_ponte("fill_repeater", {"payload": payload}, timeout=timeout)
        dt = time.time() - t0
        if not isinstance(res_ponte, dict):
            return {
                "ok": False,
                "erro": (
                    f"TIMEOUT ({timeout:.0f}s) aguardando resposta do navegador ao preencher '{nome_exibicao}'. "
                    f"(Polls recebidos até agora: {self.qtd_polls_recebidos})"
                ),
            }
        res_ponte["_tempo_s"] = round(dt, 2)
        return res_ponte


# =============================================================================
# 4. CONEXÃO COM TIMER REGRESSIVO E PARADA AUTOMÁTICA COM DIAGNÓSTICO
# =============================================================================


def conectar_guia_com_debug_e_timer(
    ctrl: ControladorExecucaoOculta,
    timeout_segundos: int = TIMEOUT_CONEXAO_PADRAO,
) -> bool:
    """
    Tenta conectar à guia já aberta do Nex Privacy testando:
    1º Canal A: Porta CDP 9222 (`--remote-debugging-port=9222`).
    2º Canal B: Ponte/Extensão local na porta 8765 (`dados/extensao_background` ou Console F12).
    Se não conectar dentro de `timeout_segundos`, PARA IMEDIATAMENTE e exibe o diagnóstico completo!
    """
    log_debug("PASSO 1/3 — CONEXÃO", "Iniciando verificação dos canais de comunicação em 2º plano...", nivel="INFO")

    # 1. Testa primeiro se o Chrome já está ouvindo na porta CDP 9222
    log_debug("PASSO 1/3 — CONEXÃO", f"Testando Canal A (CDP direto em http://127.0.0.1:{ctrl.porta_cdp}/json)...", nivel="DEBUG")
    cli_cdp, msg_cdp = _tentar_conectar_via_cdp_9222(ctrl.porta_cdp)
    if cli_cdp is not None:
        ctrl.cliente_cdp = cli_cdp
        ctrl.modo_ativo = "cdp"
        etapa_dom = cli_cdp.executar_js(
            "(document.querySelector('.wizard-nav button.active') || {textContent: '1. Identificação'}).textContent.trim()"
        )
        log_debug("PASSO 1/3 — CONEXÃO", f"{msg_cdp} | Etapa ativa no DOM: '{etapa_dom}'", nivel="OK")
        return True
    log_debug("PASSO 1/3 — CONEXÃO", f"Canal A (CDP 9222) não ativo: {msg_cdp}", nivel="WARN")

    # 2. Inicia e testa o Canal B (Servidor Ponte / Extensão em 127.0.0.1:8765)
    log_debug(
        "PASSO 1/3 — CONEXÃO",
        f"Iniciando Canal B (Ponte Oculta em http://127.0.0.1:{ctrl.porta_ponte})...",
        nivel="DEBUG",
    )
    ok_srv, msg_srv = ctrl.iniciar_servidor_ponte()
    if not ok_srv:
        return abortar_com_diagnostico(
            onde_parou="PASSO 1/3 — Inicialização do Servidor Ponte Local (Porta 8765)",
            motivo_erro=msg_srv,
            detalhes_tecnicos=[
                f"Canal A (CDP 9222): {msg_cdp}",
                f"Canal B (Ponte {ctrl.porta_ponte}): {msg_srv}",
            ],
            como_resolver=[
                "Verifique se outro terminal não ficou aberto usando a porta 8765.",
                "Ou passe outra porta com `--porta 8766`.",
            ],
        )
    log_debug("PASSO 1/3 — CONEXÃO", msg_srv, nivel="OK")

    # Testa um ping rápido inicial de 2s
    res_ping = ctrl.enviar_comando_ponte("ping", timeout=2.0)
    if isinstance(res_ping, dict) and res_ping.get("ok"):
        ctrl.modo_ativo = "ponte"
        log_debug(
            "PASSO 1/3 — CONEXÃO",
            f"Guia já aberta conectada via Canal B ({res_ping.get('modo')}): "
            f"'{res_ping.get('title')}' | Etapa ativa: '{res_ping.get('etapaAtiva')}'",
            nivel="OK",
        )
        return True

    comando_console = f"fetch('http://127.0.0.1:{ctrl.porta_ponte}/bridge.js').then(r=>r.text()).then(eval)"
    print("\n" + "-" * 78)
    log_debug(
        "PASSO 1/3 — CONEXÃO",
        f"Nenhuma resposta imediata da guia do Nex Privacy. Iniciando TIMER DE {timeout_segundos}s...",
        nivel="TIMER",
    )
    print(f"   👉 Extensão automática (em chrome://extensions -> Carregar sem compactação):")
    print(f"      {PASTA_EXTENSAO_BG}")
    print(f"   👉 Ou no Console [F12] da guia aberta do Nex Privacy:")
    print(f"      {comando_console}")
    print("-" * 78, flush=True)

    t_inicio = time.time()
    ultimo_aviso = -1
    while (time.time() - t_inicio) < timeout_segundos:
        decorrido = int(time.time() - t_inicio)
        restante = max(0, timeout_segundos - decorrido)

        if decorrido != ultimo_aviso and decorrido % 3 == 0:
            ultimo_aviso = decorrido
            log_debug(
                "TIMER CONEXÃO",
                f"Aguardando sinal da guia do Nex Privacy... "
                f"Restam {restante:2d}s | Polls recebidos={ctrl.qtd_polls_recebidos} | "
                f"Bridge.js baixado={ctrl.qtd_bridge_baixados}",
                nivel="TIMER",
            )

        res_ping = ctrl.enviar_comando_ponte("ping", timeout=1.0)
        if isinstance(res_ping, dict) and res_ping.get("ok"):
            ctrl.modo_ativo = "ponte"
            log_debug(
                "PASSO 1/3 — CONEXÃO",
                f"Sinal recebido! Guia conectada: '{res_ping.get('title')}' "
                f"| Etapa ativa: '{res_ping.get('etapaAtiva')}' | Modo: {res_ping.get('modo')}",
                nivel="OK",
            )
            return True

        # Também re-testa CDP 9222 caso o usuário tenha aberto com porta 9222 nesse meio tempo
        cli_cdp, _ = _tentar_conectar_via_cdp_9222(ctrl.porta_cdp)
        if cli_cdp is not None:
            ctrl.cliente_cdp = cli_cdp
            ctrl.modo_ativo = "cdp"
            log_debug("PASSO 1/3 — CONEXÃO", "Conectado via Canal A (CDP 9222)!", nivel="OK")
            return True

    # SE CHEGOU AQUI, O TIMER EXPIROU! PARA IMEDIATAMENTE E EXPLICA O ERRO!
    detalhes = [
        f"Tempo limite esgotado ({timeout_segundos}s) esperando resposta da guia '{URL_PADRAO_NEX_PRIVACY}'.",
        f"Canal A (CDP porta {ctrl.porta_cdp}): {msg_cdp}.",
        f"Canal B (Ponte porta {ctrl.porta_ponte}): Polls recebidos = {ctrl.qtd_polls_recebidos}, Downloads de /bridge.js = {ctrl.qtd_bridge_baixados}.",
    ]
    if ctrl.qtd_polls_recebidos == 0 and ctrl.qtd_bridge_baixados == 0:
        detalhes.append(
            "CAUSA DETECTADA: O navegador Google Chrome aberto está isolado (nem a extensão de "
            "`dados/extensao_background` foi carregada em `chrome://extensions`, nem o comando do Console F12 foi executado na guia)."
        )
    elif ctrl.qtd_polls_recebidos > 0:
        detalhes.append(
            "CAUSA DETECTADA: A extensão respondeu ao poll, mas NÃO encontrou nenhuma aba aberta com a URL "
            "`https://parana.nexprivacy.com/admin/institutional-documents/...` neste navegador."
        )

    return abortar_com_diagnostico(
        onde_parou="PASSO 1/3 — Conexão com a guia já aberta do Nex Privacy",
        motivo_erro=f"A guia do sistema não respondeu dentro do timer de {timeout_segundos} segundos.",
        detalhes_tecnicos=detalhes,
        como_resolver=[
            f"Abra `chrome://extensions` no seu Chrome, ative 'Modo do desenvolvedor', clique em 'Carregar sem compactação' e selecione a pasta:\n      {PASTA_EXTENSAO_BG}",
            "Certifique-se de que a guia `https://parana.nexprivacy.com/admin/institutional-documents/100/edit` está aberta nesse mesmo Chrome.",
            "Se preferir usar o Console (F12) na guia aberta sem instalar a extensão, lembre-se de que o Chrome pede para digitar `allow pasting` na 1ª vez antes de colar o comando `fetch(...)`.",
        ],
    )


# =============================================================================
# 5. LEITURA E VALIDAÇÃO DA PLANILHA COM DEBUG E TIMER
# =============================================================================


def _pontuar_arquivo_xlsx_pdtic(caminho: Path) -> Tuple[int, List[str]]:
    """Retorna `(score, lista_de_abas)` do arquivo `.xlsx`."""
    try:
        wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
        abas_orig = list(wb.sheetnames)
        nomes = [remover_acentos(n) for n in abas_orig]
        wb.close()
        score = 0
        for padrao in ("3.1", "3.2", "4.1", "4.2", "4.3", "5.1", "5.3", "SWOT", "METAS"):
            if any(padrao in n for n in nomes):
                score += 10
        return score, abas_orig
    except Exception:
        return -1, []


def localizar_arquivo_planilha_xlsx(caminho_informado: Optional[str] = None) -> Optional[Path]:
    if caminho_informado:
        p = Path(caminho_informado).expanduser().resolve()
        if p.is_file():
            return p

    # Prioridade máxima: arquivo padrão na raiz do projeto ou em dados/
    for base in [Path.cwd(), DADOS_DIR]:
        direto = base / NOME_ARQUIVO_EXCEL_PADRAO
        if direto.is_file():
            return direto.resolve()

    candidatos: List[Tuple[int, float, Path]] = []
    pastas_busca = [
        Path.cwd(),
        DADOS_DIR,
        Path.home() / "Downloads",
    ]
    vistos = set()
    for pasta in pastas_busca:
        if not pasta.is_dir():
            continue
        for arq in pasta.glob("*.xlsx"):
            if arq.name.startswith("~$") or arq.name.lower() == "exemplo.xlsx":
                continue
            resolvido = arq.resolve()
            if resolvido in vistos:
                continue
            vistos.add(resolvido)
            score, _ = _pontuar_arquivo_xlsx_pdtic(resolvido)
            nome_norm = remover_acentos(arq.name)
            if any(termo in nome_norm for termo in ("TECPAR", "CAMPOS NECESSARIOS", "CONSOLIDADO", "SESP")):
                score += 50
            if score > 0:
                candidatos.append((score, resolvido.stat().st_mtime, resolvido))

    if not candidatos:
        return None

    candidatos.sort(key=lambda item: (item[0], item[1]), reverse=True)
    return candidatos[0][2]


def obter_planilha_com_debug_e_timer(
    ctrl: ControladorExecucaoOculta,
    caminho_informado: Optional[str] = None,
) -> Optional[Path]:
    """
    Obtém e valida a planilha `'Campos Necessários para Sistema da Tecpar - SESP'`
    mostrando no `[DEBUG]` cada tentativa e o tempo gasto.
    """
    log_debug(
        "PASSO 2/3 — PLANILHA",
        f"Procurando planilha '{NOME_PADRAO_PLANILHA_DRIVE}' em `dados/` ou `Downloads`...",
        nivel="DEBUG",
    )
    local_xlsx = localizar_arquivo_planilha_xlsx(caminho_informado)
    if local_xlsx is not None:
        score, abas = _pontuar_arquivo_xlsx_pdtic(local_xlsx)
        log_debug(
            "PASSO 2/3 — PLANILHA",
            f"Arquivo `.xlsx` local encontrado: '{local_xlsx}' (score={score}, abas={abas})",
            nivel="OK",
        )
        return local_xlsx

    info_hist = descobrir_urls_no_historico_chrome(NOME_PADRAO_PLANILHA_DRIVE)
    sheet_id = str(info_hist.get("sheet_id") or ID_PADRAO_PLANILHA_DRIVE)
    export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    destino = DADOS_DIR / "planilha_pdtic_google_sheets.xlsx"

    log_debug(
        "PASSO 2/3 — PLANILHA",
        f"Nenhum `.xlsx` local encontrado. Tentando download HTTP direto em memória (timeout=8s): {export_url}",
        nivel="DEBUG",
    )
    try:
        req = urllib.request.Request(export_url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=8.0) as resp:
            ct = str(resp.headers.get("Content-Type") or "")
            conteudo = resp.read()
            if "text/html" not in ct.lower() and len(conteudo) > 1000:
                DADOS_DIR.mkdir(parents=True, exist_ok=True)
                destino.write_bytes(conteudo)
                score, abas = _pontuar_arquivo_xlsx_pdtic(destino)
                if score > 0:
                    log_debug(
                        "PASSO 2/3 — PLANILHA",
                        f"Planilha baixada diretamente em memória! Abas encontradas: {abas}",
                        nivel="OK",
                    )
                    return destino
            else:
                log_debug(
                    "PASSO 2/3 — PLANILHA",
                    f"Download direto retornou Content-Type='{ct}' (requer sessão autenticada do Google Drive).",
                    nivel="WARN",
                )
    except Exception as exc:
        log_debug("PASSO 2/3 — PLANILHA", f"Download HTTP direto sem cookies não disponível ({exc}).", nivel="WARN")

    if ctrl.modo_ativo == "ponte":
        log_debug(
            "PASSO 2/3 — PLANILHA",
            "Solicitando download oculto da planilha via Service Worker da extensão no Chrome (timeout=12s)...",
            nivel="TIMER",
        )
        res_ext = ctrl.enviar_comando_ponte("export_xlsx", {"exportUrl": export_url}, timeout=12.0)
        if isinstance(res_ext, dict) and res_ext.get("ok") and res_ext.get("base64"):
            try:
                DADOS_DIR.mkdir(parents=True, exist_ok=True)
                destino.write_bytes(base64.b64decode(str(res_ext["base64"])))
                score, abas = _pontuar_arquivo_xlsx_pdtic(destino)
                if score > 0:
                    log_debug(
                        "PASSO 2/3 — PLANILHA",
                        f"Planilha '{NOME_PADRAO_PLANILHA_DRIVE}' exportada em 2º plano pelo Chrome! Abas: {abas}",
                        nivel="OK",
                    )
                    return destino
            except Exception as exc:
                log_debug("PASSO 2/3 — PLANILHA", f"Erro ao salvar `.xlsx` recebido da extensão: {exc}", nivel="ERRO")
        else:
            log_debug(
                "PASSO 2/3 — PLANILHA",
                f"A extensão/ponte não retornou o `.xlsx` (detalhe: {res_ext}).",
                nivel="WARN",
            )

    return None


def _corresponde_nome_aba_excel(nome_aba_real: str, nome_guia_esperada: str) -> int:
    n_real = remover_acentos(nome_aba_real)
    n_esp = remover_acentos(nome_guia_esperada)
    if n_real == n_esp:
        return 100
    if n_esp in n_real or (len(n_real) >= 4 and n_real in n_esp):
        return 90

    m_esp = re.search(r"\b(\d+\.\d+)\b", n_esp)
    prefixo = m_esp.group(1) if m_esp else ""
    if prefixo and prefixo in n_real:
        palavras = [p for p in re.split(r"[^A-Z0-9]+", n_esp) if len(p) >= 4]
        hits = sum(1 for p in palavras if p in n_real)
        return 80 + hits * 5

    if "SWOT" in n_esp and "SWOT" in n_real:
        return 85
    if "INVENTARIO" in n_esp and ("INVENTARIO" in n_real or "NECESSIDADE" in n_real):
        return 85
    return 0


def ler_guia_de_arquivo_xlsx(
    caminho_xlsx: Path,
    nome_guia_esperada: str,
    colunas_esperadas: List[str],
) -> List[Dict[str, Any]]:
    t0 = time.time()
    log_debug(
        nome_guia_esperada,
        f"Abrindo '{caminho_xlsx.name}' em memória para localizar a aba '{nome_guia_esperada}'...",
        nivel="DEBUG",
    )
    wb = openpyxl.load_workbook(caminho_xlsx, read_only=True, data_only=True)
    try:
        melhor_aba = None
        melhor_score = 0
        for sheet_name in wb.sheetnames:
            sc = _corresponde_nome_aba_excel(sheet_name, nome_guia_esperada)
            if sc > melhor_score:
                melhor_score = sc
                melhor_aba = sheet_name

        if not melhor_aba:
            raise RuntimeError(
                f"Aba '{nome_guia_esperada}' não encontrada em '{caminho_xlsx.name}'. "
                f"Abas disponíveis: {wb.sheetnames}"
            )

        ws = wb[melhor_aba]
        linhas_tsv: List[str] = []
        for row in ws.iter_rows(values_only=True):
            if not row or not any(cell is not None and str(cell).strip() != "" for cell in row):
                continue
            celulas: List[str] = []
            for cell in row:
                if cell is None:
                    celulas.append("")
                elif isinstance(cell, float):
                    if cell.is_integer():
                        celulas.append(str(int(cell)))
                    else:
                        celulas.append(f"{cell:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                else:
                    txt_c = re.sub(r"[\r\n\t]+", " ", str(cell)).strip()
                    celulas.append(txt_c)
            linhas_tsv.append("\t".join(celulas))

        texto_tsv = "\n".join(linhas_tsv)
        linhas_extraidas = extrair_tabela_de_texto_tsv(texto_tsv, colunas_esperadas)
        dt = time.time() - t0
        log_debug(
            nome_guia_esperada,
            f"Aba '{melhor_aba}' lida em {dt:.2f}s -> {len(linhas_extraidas)} linha(s) extraída(s) "
            f"(colunas esperadas: {colunas_esperadas})!",
            nivel="OK",
        )
        return linhas_extraidas
    finally:
        wb.close()


# =============================================================================
# 6. ROTEIROS EM 2º PLANO COM DEBUG POR ITEM E PARADA EM CASO DE ERRO
# =============================================================================


def _aplicar_limite(linhas: List[Any], limite: Optional[int]) -> List[Any]:
    if limite is not None and limite > 0 and len(linhas) > limite:
        return linhas[:limite]
    return linhas


def _preencher_repeater_coluna_unica(
    ctrl: ControladorExecucaoOculta,
    step_index: int,
    repeater_name: str,
    nome_exibicao: str,
    valores: List[str],
    limite: Optional[int] = None,
) -> bool:
    valores_lim = _aplicar_limite([v for v in valores if str(v).strip()], limite)
    if not valores_lim:
        log_debug(nome_exibicao, "Nenhum valor não-vazio encontrado na planilha para esta coluna.", nivel="WARN")
        return True

    payload = {
        "stepIndex": step_index,
        "repeaterName": repeater_name,
        "mode": "single_column",
        "items": valores_lim,
    }
    res = ctrl.preencher_repeater(payload, nome_exibicao=nome_exibicao)
    if not isinstance(res, dict) or not res.get("ok"):
        return abortar_com_diagnostico(
            onde_parou=f"Preenchimento da coluna '{nome_exibicao}' (repeater='{repeater_name}', etapa={step_index + 1})",
            motivo_erro=str(res.get("erro") if isinstance(res, dict) else res),
            detalhes_tecnicos=[
                f"Coluna: {nome_exibicao} | Seletor: [data-repeater=\"{repeater_name}\"]",
                f"Quantidade de itens enviados: {len(valores_lim)}",
                f"Retorno bruto do navegador: {res}",
            ],
            como_resolver=[
                "Verifique se a guia do Nex Privacy continua aberta na tela de edição do documento PDTIC.",
            ],
        )

    for det in (res.get("detalhes") or [])[:8]:
        log_debug(nome_exibicao, f"   ↳ {det}", nivel="DEBUG")
    log_debug(
        nome_exibicao,
        f"Concluído em {res.get('_tempo_s', 0)}s: {res.get('preenchidos')}/{len(valores_lim)} item(ns) preenchido(s) "
        f"(total de caixas no DOM: {res.get('totalLinhasDOM')})!",
        nivel="OK",
    )
    return True


def _preencher_repeater_grupo(
    ctrl: ControladorExecucaoOculta,
    step_index: int,
    repeater_name: str,
    nome_exibicao: str,
    itens_mapeados: List[Dict[str, str]],
    limite: Optional[int] = None,
) -> bool:
    itens_lim = _aplicar_limite(itens_mapeados, limite)
    if not itens_lim:
        log_debug(nome_exibicao, "Nenhuma linha encontrada na planilha para este bloco.", nivel="WARN")
        return True

    payload = {
        "stepIndex": step_index,
        "repeaterName": repeater_name,
        "mode": "multi_field",
        "items": itens_lim,
    }
    res = ctrl.preencher_repeater(payload, nome_exibicao=nome_exibicao)
    if not isinstance(res, dict) or not res.get("ok"):
        return abortar_com_diagnostico(
            onde_parou=f"Preenchimento do bloco '{nome_exibicao}' (repeater='{repeater_name}', etapa={step_index + 1})",
            motivo_erro=str(res.get("erro") if isinstance(res, dict) else res),
            detalhes_tecnicos=[
                f"Bloco: {nome_exibicao} | Seletor: [data-repeater=\"{repeater_name}\"]",
                f"Quantidade de sub-cards enviados: {len(itens_lim)}",
                f"Retorno bruto do navegador: {res}",
            ],
            como_resolver=[
                "Verifique se a guia do Nex Privacy continua aberta na tela de edição do documento PDTIC.",
            ],
        )

    for det in (res.get("detalhes") or [])[:8]:
        log_debug(nome_exibicao, f"   ↳ {det}", nivel="DEBUG")
    log_debug(
        nome_exibicao,
        f"Concluído em {res.get('_tempo_s', 0)}s: {res.get('preenchidos')}/{len(itens_lim)} sub-card(s) preenchido(s) "
        f"(total de blocos no DOM: {res.get('totalLinhasDOM')})!",
        nivel="OK",
    )
    return True


def _executar_bloco_3_1_swot(
    ctrl: ControladorExecucaoOculta,
    caminho_xlsx: Path,
    limite: Optional[int] = None,
) -> bool:
    log_debug("BLOCO 3.1", "Iniciando Etapa 3 — Bloco 3.1 ('3.1. Análise SWOT')...", nivel="INFO")
    linhas = ler_guia_de_arquivo_xlsx(caminho_xlsx, "3.1. SWOT", list(COLUNAS_PADRAO_SWOT))
    pares = [
        ("FORÇAS", "strengths"),
        ("FRAQUEZAS", "weaknesses"),
        ("OPORTUNIDADES", "opportunities"),
        ("AMEAÇAS", "threats"),
    ]
    for nome_col, rep_name in pares:
        vals = [buscar_valor_da_coluna(l, nome_col) for l in linhas]
        vals = [v for v in vals if v]
        if not _preencher_repeater_coluna_unica(ctrl, 2, rep_name, f"3.1 — {nome_col}", vals, limite):
            return False
    return True


def _executar_bloco_3_2_inventario(
    ctrl: ControladorExecucaoOculta,
    caminho_xlsx: Path,
    limite: Optional[int] = None,
) -> bool:
    log_debug("BLOCO 3.2", "Iniciando Etapa 3 — Bloco 3.2 ('3.2. Inventário de necessidades')...", nivel="INFO")
    linhas = ler_guia_de_arquivo_xlsx(
        caminho_xlsx,
        "3.2. Inventário de necessidades",
        list(COLUNAS_INVENTARIO_NECESSIDADES),
    )
    itens: List[Dict[str, str]] = []
    for l in linhas:
        item = {
            'textarea[name$="[description]"]': buscar_valor_da_coluna(l, "DESCRIÇÃO"),
            'select[name$="[priority]"]': buscar_valor_da_coluna(l, "PRIORIDADE"),
            'input[name$="[requester_area]"]': buscar_valor_da_coluna(l, "ÁREA DEMANDANTE"),
            'textarea[name$="[justification]"]': buscar_valor_da_coluna(l, "JUSTIFICATIVA"),
        }
        if any(item.values()):
            itens.append(item)
    return _preencher_repeater_grupo(
        ctrl, 2, "needs", "3.2. Inventário de necessidades", itens, limite
    )


def _executar_bloco_4_1_metas(
    ctrl: ControladorExecucaoOculta,
    caminho_xlsx: Path,
    mapeamento: Dict[str, Any],
    limite: Optional[int] = None,
) -> bool:
    log_debug("BLOCO 4.1", "Iniciando Etapa 4 — Bloco 4.1 ('4.1. Metas e Ações')...", nivel="INFO")
    linhas = ler_guia_de_arquivo_xlsx(
        caminho_xlsx,
        "4.1. Metas e Ações",
        list(COLUNAS_METAS_ACOES),
    )
    if linhas:
        salvar_dados_etapa_4_em_memoria(linhas, mapeamento)
    itens: List[Dict[str, str]] = []
    for l in linhas:
        item = {
            'input[name$="[id]"]': buscar_valor_da_coluna(l, "ID"),
            'input[name$="[strategic_objective]"]': buscar_valor_da_coluna(l, "OBJETIVO ESTRATÉGICO"),
            'input[name$="[goal]"]': buscar_valor_da_coluna(l, "META"),
            'input[name$="[main_action]"]': buscar_valor_da_coluna(l, "AÇÃO PRINCIPAL"),
            'input[name$="[indicator]"]': buscar_valor_da_coluna(l, "INDICADOR"),
            'input[name$="[deadline]"]': buscar_valor_da_coluna(l, "PRAZO"),
            'input[name$="[responsible]"]': buscar_valor_da_coluna(l, "RESPONSÁVEL"),
        }
        if any(item.values()):
            itens.append(item)
    return _preencher_repeater_grupo(
        ctrl, 3, "goals-actions", "4.1. Metas e ações", itens, limite
    )


def _executar_bloco_4_2_orcamento(
    ctrl: ControladorExecucaoOculta,
    caminho_xlsx: Path,
    limite: Optional[int] = None,
) -> bool:
    log_debug("BLOCO 4.2", "Iniciando Etapa 4 — Bloco 4.2 ('4.2. Plano Orçamentário')...", nivel="INFO")
    linhas = ler_guia_de_arquivo_xlsx(
        caminho_xlsx,
        "4.2. Plano Orçamentário",
        list(COLUNAS_PLANO_ORCAMENTARIO),
    )
    itens: List[Dict[str, str]] = []
    for l in linhas:
        item = {
            'input[name$="[action_id]"]': buscar_valor_da_coluna(l, "AÇÃO (ID)"),
            'input[name$="[cost_description]"]': buscar_valor_da_coluna(l, "DESCRIÇÃO DO CUSTO"),
            'input[name$="[funding_source]"]': buscar_valor_da_coluna(l, "FONTE DE RECURSO"),
            'input[name$="[year_1]"]': buscar_valor_da_coluna(l, "ANO 1"),
            'input[name$="[year_2]"]': buscar_valor_da_coluna(l, "ANO 2"),
            'input[name$="[total]"]': buscar_valor_da_coluna(l, "TOTAL"),
        }
        if any(item.values()):
            itens.append(item)
    return _preencher_repeater_grupo(
        ctrl, 3, "budget-plan", "4.2. Plano orçamentário e de investimentos", itens, limite
    )


def _executar_bloco_4_3_riscos(
    ctrl: ControladorExecucaoOculta,
    caminho_xlsx: Path,
    limite: Optional[int] = None,
) -> bool:
    log_debug("BLOCO 4.3", "Iniciando Etapa 4 — Bloco 4.3 ('4.3. Gestão de Riscos')...", nivel="INFO")
    linhas = ler_guia_de_arquivo_xlsx(
        caminho_xlsx,
        "4.3. Gestão de Riscos",
        list(COLUNAS_GESTAO_RISCOS),
    )
    itens: List[Dict[str, str]] = []
    for l in linhas:
        item = {
            '[name$="[risk]"]': buscar_valor_da_coluna(l, "RISCO IDENTIFICADO"),
            '[name$="[probability]"]': buscar_valor_da_coluna(l, "PROBABILIDADE"),
            '[name$="[impact]"]': buscar_valor_da_coluna(l, "IMPACTO"),
            '[name$="[level]"]': buscar_valor_da_coluna(l, "NÍVEL"),
            '[name$="[mitigation]"]': buscar_valor_da_coluna(l, "AÇÃO DE MITIGAÇÃO"),
            '[name$="[responsible]"]': buscar_valor_da_coluna(l, "RESPONSÁVEL PELO RISCO"),
        }
        if any(item.values()):
            itens.append(item)
    return _preencher_repeater_grupo(
        ctrl, 3, "risk-plan", "4.3. Plano de gestão de riscos", itens, limite
    )


def _executar_bloco_5_1_capacitacao(
    ctrl: ControladorExecucaoOculta,
    caminho_xlsx: Path,
    limite: Optional[int] = None,
) -> bool:
    log_debug("BLOCO 5.1", "Iniciando Etapa 5 — Bloco 5.1 ('5.1. Plano de Capacitação')...", nivel="INFO")
    linhas = ler_guia_de_arquivo_xlsx(
        caminho_xlsx,
        "5.1. Plano de Capacitação",
        list(COLUNAS_CAPACITACAO),
    )
    pares = [
        ("TREINAMENTOS TÉCNICOS", "technical-trainings"),
        ("AÇÕES DE RETENÇÃO DE TALENTOS", "talent-actions"),
    ]
    for nome_col, rep_name in pares:
        vals = [buscar_valor_da_coluna(l, nome_col) for l in linhas]
        vals = [v for v in vals if v]
        if not _preencher_repeater_coluna_unica(ctrl, 4, rep_name, f"5.1 — {nome_col}", vals, limite):
            return False
    return True


def _executar_bloco_5_3_acoes_especificas(
    ctrl: ControladorExecucaoOculta,
    caminho_xlsx: Path,
    limite: Optional[int] = None,
) -> bool:
    log_debug("BLOCO 5.3", "Iniciando Etapa 5 — Bloco 5.3 ('5.3. Ações Específicas')...", nivel="INFO")
    linhas = ler_guia_de_arquivo_xlsx(
        caminho_xlsx,
        "5.3. Ações Específicas",
        list(COLUNAS_ACOES_ESPECIFICAS),
    )
    pares = [
        ("AÇÕES MACRO", "macro-actions"),
        ("AÇÕES MICRO", "micro-actions"),
        ("AÇÕES EXTRAS", "extra-actions"),
    ]
    for nome_col, rep_name in pares:
        vals = [buscar_valor_da_coluna(l, nome_col) for l in linhas]
        vals = [v for v in vals if v]
        if not _preencher_repeater_coluna_unica(ctrl, 4, rep_name, f"5.3 — {nome_col}", vals, limite):
            return False
    return True


def _executar_bloco_6_kpis(
    ctrl: ControladorExecucaoOculta,
    caminho_xlsx: Optional[Path],
    mapeamento: Dict[str, Any],
    limite: Optional[int] = None,
) -> bool:
    log_debug("BLOCO 6.1", "Iniciando Etapa 6 — 'Indicadores de sucesso (KPIs)'...", nivel="INFO")
    linhas_kpis = obter_dados_etapa_4_da_memoria(mapeamento)
    if not linhas_kpis and caminho_xlsx is not None:
        linhas_41 = ler_guia_de_arquivo_xlsx(
            caminho_xlsx,
            "4.1. Metas e Ações",
            list(COLUNAS_METAS_ACOES),
        )
        linhas_kpis = salvar_dados_etapa_4_em_memoria(linhas_41, mapeamento)

    if not linhas_kpis:
        return abortar_com_diagnostico(
            onde_parou="BLOCO 6.1 — 'Indicadores de sucesso (KPIs)'",
            motivo_erro="Nenhum dado encontrado na memória da Etapa 4 (`dados/cache_etapa_4.json`) nem na planilha.",
            detalhes_tecnicos=["`obter_dados_etapa_4_da_memoria` retornou lista vazia."],
            como_resolver=["Execute primeiro o Bloco 4.1 ou coloque o arquivo `.xlsx` da planilha em `Downloads`."],
        )

    itens: List[Dict[str, str]] = []
    for l in linhas_kpis:
        ind = buscar_valor_da_coluna(l, "INDICADOR")
        form = buscar_valor_da_coluna(l, "FÓRMULA / DESCRIÇÃO")
        if ind or form:
            itens.append(
                {
                    'input[name$="[name]"]': ind,
                    'input[name$="[formula]"]': form,
                }
            )
    return _preencher_repeater_grupo(
        ctrl, 5, "kpis", "6. Indicadores de sucesso (KPIs)", itens, limite
    )


def executar_automacao_segundo_plano(
    alvo: str = "todas",
    caminho_planilha: Optional[str] = None,
    limite: Optional[int] = None,
    porta_cdp: int = PORTA_PONTE_LOCAL,
    timeout_conexao: int = TIMEOUT_CONEXAO_PADRAO,
) -> bool:
    """
    Executa a automação 100% oculta com `[DEBUG]` passo a passo e Timers de segurança
    para parar e avisar exatamente onde e por que parou caso qualquer etapa não consiga prosseguir.
    """
    chave = str(alvo or "todas").strip().lower()
    mapeamento = carregar_mapeamento() or {}

    log_debug(
        "INÍCIO",
        f"Alvo='{chave}' | Limite={limite or 'SEM LIMITE'} | Timeout Conexão={timeout_conexao}s",
        nivel="INFO",
    )

    ctrl = ControladorExecucaoOculta(porta_ponte=porta_cdp, porta_cdp=PORTA_CDP_PADRAO)
    try:
        # PASSO 1: Conectar à guia já aberta do Nex Privacy (com Timer regressivo e diagnóstico)
        if not conectar_guia_com_debug_e_timer(ctrl, timeout_segundos=timeout_conexao):
            return False

        # PASSO 2: Obter e validar a planilha 'Campos Necessários para Sistema da Tecpar - SESP'
        caminho_xlsx = obter_planilha_com_debug_e_timer(ctrl, caminho_planilha)
        if caminho_xlsx is None and chave not in ("pagina6", "p6", "6", "6.1", "kpis"):
            return abortar_com_diagnostico(
                onde_parou=f"PASSO 2/3 — Leitura da Planilha '{NOME_PADRAO_PLANILHA_DRIVE}'",
                motivo_erro="Não foi possível ler a planilha `.xlsx` localmente nem via extensão do Chrome.",
                detalhes_tecnicos=[
                    f"Buscado em: `{DADOS_DIR}`, `{Path.cwd()}` e `{Path.home() / 'Downloads'}`.",
                    f"Link Google Sheets configurado: `{URL_PADRAO_PLANILHA_DRIVE}`.",
                    f"Modo de conexão ativo: `{ctrl.modo_ativo}`.",
                ],
                como_resolver=[
                    f"No Google Sheets ('{NOME_PADRAO_PLANILHA_DRIVE}'), clique em `Arquivo -> Baixar -> Microsoft Excel (.xlsx)` (deixe na pasta `Downloads` ou `dados/`).",
                    f"Ou instale a extensão de `{PASTA_EXTENSAO_BG}` em `chrome://extensions` para que o próprio Chrome baixe a planilha oculto pelos seus cookies.",
                ],
            )

        # PASSO 3: Executar os blocos solicitados em sequência (parando imediatamente no 1º erro)
        log_debug("PASSO 3/3 — EXECUÇÃO", f"Iniciando preenchimento em 2º plano para alvo='{chave}'...", nivel="INFO")

        if chave in ("3.1", "swot"):
            return _executar_bloco_3_1_swot(ctrl, caminho_xlsx, limite)  # type: ignore[arg-type]
        if chave in ("3.2", "inventario"):
            return _executar_bloco_3_2_inventario(ctrl, caminho_xlsx, limite)  # type: ignore[arg-type]
        if chave in ("pagina3", "p3", "3"):
            return (
                _executar_bloco_3_1_swot(ctrl, caminho_xlsx, limite)  # type: ignore[arg-type]
                and _executar_bloco_3_2_inventario(ctrl, caminho_xlsx, limite)  # type: ignore[arg-type]
            )
        if chave in ("4.1", "metas-acoes"):
            return _executar_bloco_4_1_metas(ctrl, caminho_xlsx, mapeamento, limite)  # type: ignore[arg-type]
        if chave in ("4.2", "orcamento"):
            return _executar_bloco_4_2_orcamento(ctrl, caminho_xlsx, limite)  # type: ignore[arg-type]
        if chave in ("4.3", "riscos"):
            return _executar_bloco_4_3_riscos(ctrl, caminho_xlsx, limite)  # type: ignore[arg-type]
        if chave in ("pagina4", "p4", "4"):
            return (
                _executar_bloco_4_1_metas(ctrl, caminho_xlsx, mapeamento, limite)  # type: ignore[arg-type]
                and _executar_bloco_4_2_orcamento(ctrl, caminho_xlsx, limite)  # type: ignore[arg-type]
                and _executar_bloco_4_3_riscos(ctrl, caminho_xlsx, limite)  # type: ignore[arg-type]
            )
        if chave in ("5.1", "capacitacao"):
            return _executar_bloco_5_1_capacitacao(ctrl, caminho_xlsx, limite)  # type: ignore[arg-type]
        if chave in ("5.3", "acoes-especificas"):
            return _executar_bloco_5_3_acoes_especificas(ctrl, caminho_xlsx, limite)  # type: ignore[arg-type]
        if chave in ("pagina5", "p5", "5"):
            return (
                _executar_bloco_5_1_capacitacao(ctrl, caminho_xlsx, limite)  # type: ignore[arg-type]
                and _executar_bloco_5_3_acoes_especificas(ctrl, caminho_xlsx, limite)  # type: ignore[arg-type]
            )
        if chave in ("pagina6", "p6", "6", "6.1", "kpis"):
            return _executar_bloco_6_kpis(ctrl, caminho_xlsx, mapeamento, limite)

        ok = (
            _executar_bloco_3_1_swot(ctrl, caminho_xlsx, limite)  # type: ignore[arg-type]
            and _executar_bloco_3_2_inventario(ctrl, caminho_xlsx, limite)  # type: ignore[arg-type]
            and _executar_bloco_4_1_metas(ctrl, caminho_xlsx, mapeamento, limite)  # type: ignore[arg-type]
            and _executar_bloco_4_2_orcamento(ctrl, caminho_xlsx, limite)  # type: ignore[arg-type]
            and _executar_bloco_4_3_riscos(ctrl, caminho_xlsx, limite)  # type: ignore[arg-type]
            and _executar_bloco_5_1_capacitacao(ctrl, caminho_xlsx, limite)  # type: ignore[arg-type]
            and _executar_bloco_5_3_acoes_especificas(ctrl, caminho_xlsx, limite)  # type: ignore[arg-type]
            and _executar_bloco_6_kpis(ctrl, caminho_xlsx, mapeamento, limite)
        )
        if ok:
            log_debug(
                "FINALIZADO",
                "🎉 Todas as etapas (3, 4, 5 e 6) foram preenchidas em 2º plano com sucesso!",
                nivel="OK",
            )
        return ok
    finally:
        ctrl.fechar()

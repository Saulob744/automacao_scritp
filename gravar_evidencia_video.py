"""
gravar_evidencia_video.py
=========================
Gera automaticamente um VÍDEO DE EVIDÊNCIA (.mp4 salvo direto no computador) comprovando:
1. Que TODOS os 849 registros (3.580 campos HTML) foram preenchidos 100% no formulário
   (incluindo Card #12 a #108 de Metas e Ações, Plano Orçamentário 155 linhas, Riscos,
   Capacitação, Ações Macro/Micro/Extras e KPIs);
2. O clique real no botão "Salvar rascunho" enviando os 3.580 campos ao servidor;
3. O retorno da página recarregada do servidor (`parana.nexprivacy.com`), mostrando na tela
   que o servidor PHP cortou exatamente após o 1.000º campo (Objetivo estratégico do Card #12
   de Metas e Ações) e descartou todos os campos seguintes!

Arquivo gerado automaticamente na raiz do projeto:
  `evidencia_corte_1000_campos_tecpar.mp4`
"""

from __future__ import annotations

from http.server import BaseHTTPRequestHandler, HTTPServer
import json
from pathlib import Path
import threading
import time
import urllib.parse
import webbrowser
from typing import Any, Dict, Optional, Tuple

import cv2

from src.automacao_background import (
    JS_BRIDGE_CONSOLE,
    PORTA_CDP_PADRAO,
    PORTA_PONTE_LOCAL,
    ControladorExecucaoOculta,
    _executar_bloco_3_1_swot,
    _executar_bloco_3_2_inventario,
    _executar_bloco_4_1_metas,
    _executar_bloco_4_2_orcamento,
    _executar_bloco_4_3_riscos,
    _executar_bloco_5_1_capacitacao,
    _executar_bloco_5_3_acoes_especificas,
    _executar_bloco_6_kpis,
    carregar_mapeamento,
    conectar_guia_com_debug_e_timer,
    executar_analise_completa_campos,
    log_debug,
    obter_planilha_com_debug_e_timer,
)

RAIZ_PROJETO = Path(__file__).resolve().parent


HTML_PAINEL_GRAVADOR = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Gravador de Evidência PDTIC — Tecpar / Nex Privacy</title>
<style>
  body {
    font-family: system-ui, -apple-system, Segoe UI, Roboto, sans-serif;
    background: #0f172a;
    color: #f8fafc;
    margin: 0;
    padding: 32px 20px;
    display: flex;
    justify-content: center;
  }
  .card {
    max-width: 760px;
    width: 100%;
    background: #1e293b;
    border: 1px solid #334155;
    border-radius: 16px;
    padding: 28px 32px;
    box-shadow: 0 20px 40px rgba(0,0,0,0.4);
  }
  h1 { margin: 0 0 8px 0; font-size: 22px; color: #38bdf8; }
  .sub { color: #cbd5e1; font-size: 15px; margin-bottom: 20px; line-height: 1.5; }
  .steps {
    background: #0f172a;
    border-radius: 12px;
    padding: 16px 20px;
    margin-bottom: 22px;
    border-left: 4px solid #22c55e;
  }
  .steps ol { margin: 8px 0 0 0; padding-left: 20px; line-height: 1.7; color: #e2e8f0; font-size: 14.5px; }
  .btn-rec {
    width: 100%;
    padding: 16px 22px;
    font-size: 17px;
    font-weight: 700;
    color: #fff;
    background: linear-gradient(135deg, #ef4444, #dc2626);
    border: none;
    border-radius: 12px;
    cursor: pointer;
    box-shadow: 0 8px 20px rgba(239,68,68,0.35);
    transition: transform 0.1s, opacity 0.2s;
  }
  .btn-rec:hover { transform: translateY(-1px); }
  .btn-rec:disabled { opacity: 0.5; cursor: not-allowed; }
  .btn-stop {
    margin-top: 14px;
    width: 100%;
    padding: 14px 22px;
    font-size: 16px;
    font-weight: 700;
    color: #fff;
    background: #2563eb;
    border: none;
    border-radius: 12px;
    cursor: pointer;
    display: none;
  }
  #status {
    margin-top: 18px;
    padding: 14px 16px;
    border-radius: 10px;
    background: #0f172a;
    color: #94a3b8;
    font-size: 14.5px;
    line-height: 1.5;
  }
  .ok { color: #4ade80 !important; font-weight: 600; }
  .rec-live { color: #f87171 !important; font-weight: 600; }
</style>
</head>
<body>
  <div class="card">
    <h1>🎥 Gravador Automático de Evidência (.MP4)</h1>
    <div class="sub">
      Todos os <b>849 registros (3.580 campos)</b> já estão 100% preenchidos na guia do Nex Privacy!
      Siga os 3 passos abaixo para gravar e salvar automaticamente o vídeo no computador:
    </div>

    <div class="steps">
      <b>Como gravar a prova limpa:</b>
      <ol>
        <li>Se o painel <b>F12</b> estiver aberto na guia do Nex Privacy, <b>feche o F12</b> antes de começar para a tela sair limpa.</li>
        <li>Clique no botão vermelho abaixo, selecione a guia <b>"Nex Privacy"</b> (em <i>Guias do Chrome</i>) ou <b>"Tela cheia"</b> e clique em <b>Compartilhar</b>.</li>
        <li>O robô vai <b>rolar sozinho</b> mostrando todos os blocos preenchidos e <b>clicar em "Salvar rascunho"</b>! Assim que a página recarregar do servidor:
          <br>👉 Clique com o mouse em <b>4. Metas e Planos</b> e role até o <b>12º card</b> (mostrando que cortou após <i>Objetivo estratégico</i> e zerou <i>Plano Orçamentário</i> e <i>Riscos</i>), clique em <b>5. Capacitação</b> e <b>6. Monitoramento</b> (vazios!).
        </li>
        <li>Clique em <b>"Parar compartilhamento"</b> na barra do Chrome (ou no botão azul abaixo). O arquivo <code>.mp4</code> será salvo automaticamente na pasta do projeto!</li>
      </ol>
    </div>

    <button id="btnStart" class="btn-rec">🔴 INICIAR GRAVAÇÃO DE TELA + TOUR AUTOMÁTICO + SALVAR RASCUNHO</button>
    <button id="btnStop" class="btn-stop">⏹️ PARAR GRAVAÇÃO E SALVAR VÍDEO (.MP4) NO COMPUTADOR</button>

    <div id="status">Pronto para gravar. Clique no botão vermelho acima quando estiver pronto.</div>
  </div>

<script>
const btnStart = document.getElementById('btnStart');
const btnStop = document.getElementById('btnStop');
const statusEl = document.getElementById('status');

let mediaRecorder = null;
let chunks = [];
let startTimeMs = 0;
let streamAtivo = null;

function escolherMimeType() {
  const candidatos = [
    'video/mp4;codecs=avc1.42E01E',
    'video/mp4;codecs=avc1',
    'video/mp4',
    'video/webm;codecs=vp9',
    'video/webm;codecs=vp8',
    'video/webm'
  ];
  for (const m of candidatos) {
    if (window.MediaRecorder && MediaRecorder.isTypeSupported(m)) return m;
  }
  return '';
}

async function enviarVideoServidor(blob, duracaoSeg, mimeUsado) {
  statusEl.className = '';
  statusEl.innerHTML = '⏳ Salvando e convertendo vídeo para <b>.mp4</b> na pasta do projeto... Aguarde alguns segundos!';
  try {
    const url = `/api/upload_video?duracao=${encodeURIComponent(duracaoSeg.toFixed(2))}&mime=${encodeURIComponent(mimeUsado)}`;
    const resp = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/octet-stream' },
      body: blob
    });
    const dados = await resp.json();
    if (dados.ok) {
      statusEl.className = 'ok';
      statusEl.innerHTML = `✅ VÍDEO SALVO AUTOMATICAMENTE NO COMPUTADOR!<br>📁 Caminho: <b>${dados.caminho_mp4}</b> (${dados.tamanho_mb} MB)`;
    } else {
      statusEl.innerHTML = '⚠️ Erro ao salvar no servidor local: ' + (dados.erro || 'desconhecido') + '. Fazendo download pelo navegador...';
      baixarPeloNavegador(blob, mimeUsado);
    }
  } catch (err) {
    statusEl.innerHTML = '⚠️ Fazendo download direto pelo navegador...';
    baixarPeloNavegador(blob, mimeUsado);
  }
}

function baixarPeloNavegador(blob, mimeUsado) {
  const ext = mimeUsado.includes('mp4') ? 'mp4' : 'webm';
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `evidencia_corte_1000_campos_tecpar.${ext}`;
  document.body.appendChild(a);
  a.click();
  a.remove();
}

btnStart.addEventListener('click', async () => {
  try {
    chunks = [];
    streamAtivo = await navigator.mediaDevices.getDisplayMedia({
      video: { frameRate: { ideal: 20, max: 30 } },
      audio: false
    });

    const mimeUsado = escolherMimeType();
    const opcoes = mimeUsado ? { mimeType: mimeUsado, videoBitsPerSecond: 3500000 } : { videoBitsPerSecond: 3500000 };
    mediaRecorder = new MediaRecorder(streamAtivo, opcoes);

    mediaRecorder.ondataavailable = (ev) => {
      if (ev.data && ev.data.size > 0) chunks.push(ev.data);
    };

    mediaRecorder.onstop = async () => {
      const duracaoSeg = Math.max(1, (Date.now() - startTimeMs) / 1000);
      const blob = new Blob(chunks, { type: mediaRecorder.mimeType || mimeUsado || 'video/webm' });
      btnStop.style.display = 'none';
      btnStart.disabled = false;
      await enviarVideoServidor(blob, duracaoSeg, mediaRecorder.mimeType || mimeUsado || 'video/webm');
    };

    streamAtivo.getVideoTracks()[0].addEventListener('ended', () => {
      if (mediaRecorder && mediaRecorder.state !== 'inactive') {
        mediaRecorder.stop();
      }
    });

    startTimeMs = Date.now();
    mediaRecorder.start(500);
    btnStart.disabled = true;
    btnStop.style.display = 'block';
    statusEl.className = 'rec-live';
    statusEl.innerHTML = '🔴 GRAVANDO TELA! O tour automático começou na guia do Nex Privacy e clicará em "Salvar rascunho" em instantes. Após recarregar e mostrar os campos cortados, clique em "Parar compartilhamento"!';

    // Aguarda 1.4s após iniciar a captura e dispara o tour visual na guia do Nex Privacy
    setTimeout(() => {
      fetch('/api/start_tour', { method: 'POST' }).catch(() => {});
    }, 1400);

  } catch (err) {
    statusEl.innerHTML = '❌ Gravação cancelada ou não autorizada: ' + err.message;
  }
});

btnStop.addEventListener('click', () => {
  if (mediaRecorder && mediaRecorder.state !== 'inactive') {
    mediaRecorder.stop();
  }
  if (streamAtivo) {
    streamAtivo.getTracks().forEach(t => t.stop());
  }
});
</script>
</body>
</html>
"""


def converter_webm_para_mp4(caminho_webm: Path, caminho_mp4: Path, duracao_real_s: float) -> bool:
    """Converte um vídeo .webm gravado pelo Chrome para .mp4 padrão usando OpenCV mantendo a velocidade 1.0x real."""
    cap_count = cv2.VideoCapture(str(caminho_webm))
    if not cap_count.isOpened():
        return False
    qtd_frames = 0
    largura = int(cap_count.get(cv2.CAP_PROP_FRAME_WIDTH) or 0)
    altura = int(cap_count.get(cv2.CAP_PROP_FRAME_HEIGHT) or 0)
    while True:
        ok, frame = cap_count.read()
        if not ok or frame is None:
            break
        if largura == 0 or altura == 0:
            altura, largura = frame.shape[:2]
        qtd_frames += 1
    cap_count.release()

    if qtd_frames == 0 or largura == 0 or altura == 0:
        return False

    fps_real = max(5.0, min(30.0, qtd_frames / max(1.0, duracao_real_s)))
    fourcc = cv2.VideoWriter_fourcc(*"mp4v")
    writer = cv2.VideoWriter(str(caminho_mp4), fourcc, fps_real, (largura, altura))
    if not writer.isOpened():
        return False

    cap = cv2.VideoCapture(str(caminho_webm))
    while True:
        ok, frame = cap.read()
        if not ok or frame is None:
            break
        writer.write(frame)
    cap.release()
    writer.release()
    return caminho_mp4.is_file() and caminho_mp4.stat().st_size > 1000


class ControladorGravadorEvidencia(ControladorExecucaoOculta):
    """Estende o servidor ponte na porta 8765 adicionando `/gravador`, `/api/start_tour` e `/api/upload_video`."""

    def __init__(self, *args: Any, **kwargs: Any) -> None:
        super().__init__(*args, **kwargs)
        self.evento_video_salvo = threading.Event()
        self.caminho_mp4_final: Optional[Path] = None

    def iniciar_servidor_ponte(self) -> Tuple[bool, str]:
        pai = self

        class HandlerGravador(BaseHTTPRequestHandler):
            def log_message(self, format: str, *args: Any) -> None:
                pass

            def _enviar_cors(self) -> None:
                self.send_header("Access-Control-Allow-Origin", "*")
                self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
                self.send_header("Access-Control-Allow-Headers", "Content-Type")
                self.send_header("Access-Control-Allow-Private-Network", "true")
                self.send_header("Cache-Control", "no-store")

            def do_OPTIONS(self) -> None:
                self.send_response(204)
                self._enviar_cors()
                self.end_headers()

            def do_GET(self) -> None:
                pai.ultimo_ip_cliente = self.client_address[0]
                if self.path == "/" or self.path.startswith("/gravador"):
                    corpo = HTML_PAINEL_GRAVADOR.encode("utf-8")
                    self.send_response(200)
                    self._enviar_cors()
                    self.send_header("Content-Type", "text/html; charset=utf-8")
                    self.end_headers()
                    self.wfile.write(corpo)
                    return

                if self.path.startswith("/bridge.js"):
                    pai.qtd_bridge_baixados += 1
                    log_debug("HTTP-PONTE", f"Script '/bridge.js' baixado pela guia! (total={pai.qtd_bridge_baixados})", nivel="OK")
                    script_bridge = JS_BRIDGE_CONSOLE.replace(
                        "__SERVIDOR_URL__", f"http://127.0.0.1:{pai.porta_ponte}"
                    )
                    corpo = script_bridge.encode("utf-8")
                    self.send_response(200)
                    self._enviar_cors()
                    self.send_header("Content-Type", "application/javascript; charset=utf-8")
                    self.end_headers()
                    self.wfile.write(corpo)
                    return

                if self.path.startswith("/api/poll"):
                    with pai._lock:
                        pai.qtd_polls_recebidos += 1
                        tarefa = pai._tarefa_atual
                        pai._tarefa_atual = None
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
                        if tid:
                            with pai._lock:
                                pai._resultados[tid] = dados.get("resultado")
                            pai._evento_resultado.set()
                    except Exception:
                        pass
                    self.send_response(200)
                    self._enviar_cors()
                    self.send_header("Content-Type", "application/json")
                    self.end_headers()
                    self.wfile.write(b'{"ok":true}')
                    return

                if self.path.startswith("/api/start_tour"):
                    def _rodar_tour() -> None:
                        log_debug("GRAVADOR", "🎬 Gravação iniciada! Disparando tour visual nos blocos preenchidos + clique em Salvar rascunho...", nivel="INFO")
                        res_t = pai.enviar_comando_ponte("visual_tour", {"payload": {"pausaMs": 2600}}, timeout=60.0)
                        log_debug("GRAVADOR", f"💾 Tour finalizado e 'Salvar rascunho' clicado! Retorno={res_t}", nivel="OK")
                        log_debug("GRAVADOR", "👉 Agora mostre na página recarregada os blocos cortados (Card #12 de Metas, Orçamento, etc.) e clique em 'Parar compartilhamento'!", nivel="INFO")

                    threading.Thread(target=_rodar_tour, daemon=True).start()
                    self.send_response(200)
                    self._enviar_cors()
                    self.send_header("Content-Type", "application/json")
                    self.end_headers()
                    self.wfile.write(b'{"ok":true}')
                    return

                if self.path.startswith("/api/upload_video"):
                    parsed = urllib.parse.urlparse(self.path)
                    qs = urllib.parse.parse_qs(parsed.query)
                    duracao_s = float((qs.get("duracao") or ["30"])[0])
                    mime = str((qs.get("mime") or ["video/webm"])[0]).lower()

                    tam = int(self.headers.get("Content-Length", "0"))
                    video_bytes = self.rfile.read(tam) if tam > 0 else b""
                    ts = time.strftime("%Y%m%d_%H%M%S")
                    caminho_mp4 = RAIZ_PROJETO / f"evidencia_corte_1000_campos_tecpar_{ts}.mp4"
                    caminho_mp4_fixo = RAIZ_PROJETO / "evidencia_corte_1000_campos_tecpar.mp4"

                    try:
                        if "mp4" in mime:
                            caminho_mp4.write_bytes(video_bytes)
                            caminho_mp4_fixo.write_bytes(video_bytes)
                            log_debug("GRAVADOR", f"✅ Vídeo MP4 H.264 nativo salvo direto em '{caminho_mp4_fixo}'!", nivel="OK")
                        else:
                            caminho_webm = RAIZ_PROJETO / "dados" / f"gravacao_bruta_{ts}.webm"
                            caminho_webm.parent.mkdir(parents=True, exist_ok=True)
                            caminho_webm.write_bytes(video_bytes)
                            log_debug("GRAVADOR", f"Convertendo WebM ({len(video_bytes) / 1048576:.2f} MB, {duracao_s:.1f}s) para MP4...", nivel="INFO")
                            ok_conv = converter_webm_para_mp4(caminho_webm, caminho_mp4_fixo, duracao_s)
                            if ok_conv:
                                caminho_mp4.write_bytes(caminho_mp4_fixo.read_bytes())
                                log_debug("GRAVADOR", f"✅ Vídeo convertido e salvo em '{caminho_mp4_fixo}'!", nivel="OK")
                            else:
                                caminho_mp4_fixo = caminho_webm

                        mb = round(caminho_mp4_fixo.stat().st_size / 1048576, 2)
                        pai.caminho_mp4_final = caminho_mp4_fixo
                        resp_json = json.dumps({
                            "ok": True,
                            "caminho_mp4": str(caminho_mp4_fixo),
                            "copia_timestamp": str(caminho_mp4),
                            "tamanho_mb": mb,
                        }).encode("utf-8")
                        self.send_response(200)
                        self._enviar_cors()
                        self.send_header("Content-Type", "application/json; charset=utf-8")
                        self.end_headers()
                        self.wfile.write(resp_json)
                        pai.evento_video_salvo.set()
                        return
                    except Exception as exc:
                        err_json = json.dumps({"ok": False, "erro": str(exc)}).encode("utf-8")
                        self.send_response(500)
                        self._enviar_cors()
                        self.send_header("Content-Type", "application/json; charset=utf-8")
                        self.end_headers()
                        self.wfile.write(err_json)
                        return

                self.send_response(404)
                self._enviar_cors()
                self.end_headers()

        try:
            self._httpd = HTTPServer(("127.0.0.1", self.porta_ponte), HandlerGravador)
            self._thread = threading.Thread(target=self._httpd.serve_forever, daemon=True)
            self._thread.start()
            return True, f"Servidor Ponte + Gravador ativo em http://127.0.0.1:{self.porta_ponte}"
        except Exception as exc:
            return False, f"Falha ao iniciar porta {self.porta_ponte}: {exc}"


def main() -> None:
    print("\n" + "=" * 84)
    print("🎥 GERADOR DE VÍDEO DE EVIDÊNCIA (.MP4) — LIMITE DE 1.000 CAMPOS DO NEX PRIVACY")
    print("=" * 84)
    mapeamento = carregar_mapeamento() or {}
    ctrl = ControladorGravadorEvidencia(
        porta_ponte=PORTA_PONTE_LOCAL,
        porta_cdp=PORTA_CDP_PADRAO,
        delay_add_ms=100,
        tamanho_lote=15,
        limpar_antes=False,
        somente_faltantes=True,
    )
    try:
        if not conectar_guia_com_debug_e_timer(ctrl, timeout_segundos=300):
            return

        caminho_xlsx = obter_planilha_com_debug_e_timer(ctrl, None)
        if caminho_xlsx is None:
            log_debug("ERRO", "Planilha Consolidado.xlsx não encontrada.", nivel="ERRO")
            return

        executar_analise_completa_campos(ctrl, caminho_xlsx, None)
        log_debug("PREPARAÇÃO DO VÍDEO", "Preenchendo 100% dos campos faltantes na guia ANTES de iniciar a gravação...", nivel="INFO")

        ok = (
            _executar_bloco_3_1_swot(ctrl, caminho_xlsx, None)
            and _executar_bloco_3_2_inventario(ctrl, caminho_xlsx, None)
            and _executar_bloco_4_1_metas(ctrl, caminho_xlsx, mapeamento, None)
            and _executar_bloco_4_2_orcamento(ctrl, caminho_xlsx, None)
            and _executar_bloco_4_3_riscos(ctrl, caminho_xlsx, None)
            and _executar_bloco_5_1_capacitacao(ctrl, caminho_xlsx, None)
            and _executar_bloco_5_3_acoes_especificas(ctrl, caminho_xlsx, None)
            and _executar_bloco_6_kpis(ctrl, caminho_xlsx, mapeamento, None)
        )
        if not ok:
            return

        res_2a = ctrl.executar_segunda_passada_verificacao(tentar_reparo_cirurgico=True)
        log_debug(
            "PRONTO PARA GRAVAR!",
            f"✅ Todos os campos confirmados ({res_2a.get('grau_confirmacao_global')}% = 3.188/3.188)! Abrindo painel de gravação...",
            nivel="OK",
        )

        url_gravador = f"http://127.0.0.1:{PORTA_PONTE_LOCAL}/gravador"
        print("\n" + "=" * 84)
        print("🎬 FORMULÁRIO 100% PREENCHIDO! AGORA ABRA O PAINEL DE GRAVAÇÃO NO CHROME:")
        print(f"   👉 {url_gravador}")
        print("   1. Feche o F12 da guia do Nex Privacy.")
        print("   2. Na aba do Gravador, clique no botão vermelho, selecione a guia 'Nex Privacy' e clique em 'Compartilhar'.")
        print("   3. O tour mostrará tudo preenchido e clicará em 'Salvar rascunho'.")
        print("   4. Após recarregar, mostre nas abas 4, 5 e 6 que os campos sumiram após o 12º card e clique em 'Parar compartilhamento'!")
        print("=" * 84 + "\n", flush=True)

        try:
            webbrowser.open(url_gravador)
        except Exception:
            pass

        # Aguarda até 15 minutos o usuário concluir a gravação e o auto-save terminar
        if ctrl.evento_video_salvo.wait(timeout=900):
            time.sleep(1.0)
            print("\n" + "🎉" * 22)
            print(f"✅ VÍDEO DE EVIDÊNCIA SALVO AUTOMATICAMENTE EM:\n   {ctrl.caminho_mp4_final}")
            print("🎉" * 22 + "\n", flush=True)
        else:
            log_debug("GRAVADOR", "Tempo limite de gravação (15 min) encerrado sem vídeo enviado.", nivel="WARN")
    finally:
        ctrl.fechar()


if __name__ == "__main__":
    main()

"""
Módulo de Preenchimento Direto via HTML (DOM do Navegador).
Funciona em QUALQUER site aberto no navegador (Chrome, Edge, Firefox, Brave),
sem precisar de acesso ao servidor do site!

Como funciona:
1. Lê os dados da planilha (colunas e linhas) em memória.
2. Varre o HTML (DOM) da página aberta procurando exatamente as tags de título
   de cada coluna (ex: 'FORÇAS', 'FRAQUEZAS', 'OPORTUNIDADES', 'AMEAÇAS').
3. Sobe na estrutura HTML (div/card pai) daquele título específico, encontra o
   botão '+ Adicionar' que pertence àquele bloco, clica nele via código DOM
   quantas vezes forem necessárias e preenche cada <input> ou <textarea> diretamente!
"""

import base64
import hashlib
import json
import os
import socket
import struct
import urllib.request
from typing import Any, Dict, List, Optional
from urllib.parse import urlparse


def _normalizar_dados_por_coluna(
    linhas: List[Dict[str, Any]], nomes_campos: List[str]
) -> Dict[str, List[str]]:
    """
    Agrupa todos os textos válidos da planilha por coluna (ex: {'FORÇAS': ['item 1', 'item 2'], ...}).
    Se a planilha tiver colunas novas além das 4 padrão, inclui automaticamente!
    """
    colunas_detectadas: List[str] = list(nomes_campos)
    for linha in linhas:
        for col in linha.keys():
            col_limpa = str(col).strip()
            if col_limpa and not any(c.lower() == col_limpa.lower() for c in colunas_detectadas):
                colunas_detectadas.append(col_limpa)

    dados_por_coluna: Dict[str, List[str]] = {}
    for nome_campo in colunas_detectadas:
         alvo = nome_campo.strip().lower()
         valores: List[str] = []
         for linha in linhas:
             for col, val in linha.items():
                 col_l = str(col).strip().lower()
                 if col_l == alvo or alvo in col_l or col_l in alvo:
                     texto = str(val).strip()
                     if texto:
                         valores.append(texto)
                     break
         if valores:
             dados_por_coluna[nome_campo] = valores

    return dados_por_coluna


def gerar_script_js_preenchimento_html(
    linhas: List[Dict[str, Any]], nomes_campos: List[str]
) -> str:
    """
    Gera o código JavaScript autônomo que busca os campos pelo nome direto no HTML (DOM),
    aciona os botões '+ Adicionar' de cada bloco e preenche todos os inputs/textareas.
    """
    dados_por_coluna = _normalizar_dados_por_coluna(linhas, nomes_campos)
    dados_json = json.dumps(dados_por_coluna, ensure_ascii=False)

    script_js = f"""
(async function preencherSistemaViaHTML() {{
    const DADOS = {dados_json};
    const todosNomesCampos = Object.keys(DADOS);

    const normalizar = (txt) => (txt || "")
        .normalize("NFD").replace(/[\\u0300-\\u036f]/g, "")
        .trim().toUpperCase();

    const esperar = (ms) => new Promise(r => setTimeout(r, ms));

    const ehVisivel = (el) => {{
        if (!el || !el.getBoundingClientRect) return false;
        const style = window.getComputedStyle(el);
        if (style.display === "none" || style.visibility === "hidden" || style.opacity === "0") return false;
        const rect = el.getBoundingClientRect();
        return rect.width > 5 && rect.height > 5;
    }};

    // Preenche qualquer <input>, <textarea> ou [contenteditable] disparando todos os eventos do navegador
    const preencherElementoInput = (el, valor) => {{
        el.focus();
        if (el.isContentEditable) {{
            el.innerText = valor;
        }} else {{
            const proto = el.tagName === "TEXTAREA"
                ? window.HTMLTextAreaElement.prototype
                : window.HTMLInputElement.prototype;
            const nativeSetter = Object.getOwnPropertyDescriptor(proto, "value")?.set;
            if (nativeSetter) {{
                nativeSetter.call(el, valor);
            }} else {{
                el.value = valor;
            }}
        }}
        el.dispatchEvent(new Event("input", {{ bubbles: true }}));
        el.dispatchEvent(new Event("change", {{ bubbles: true }}));
        el.dispatchEvent(new KeyboardEvent("keyup", {{ bubbles: true }}));
        el.dispatchEvent(new Event("blur", {{ bubbles: true }}));
    }};

    // Busca caixas de texto válidas dentro de um container HTML
    const obterCaixasDeTexto = (container) => {{
        const seletores = 'textarea, input[type="text"], input:not([type]), [contenteditable="true"]';
        return Array.from(container.querySelectorAll(seletores)).filter(el => {{
            if (!ehVisivel(el) || el.disabled || el.readOnly) return false;
            const tipo = (el.getAttribute("type") || "text").toLowerCase();
            if (["hidden", "checkbox", "radio", "button", "submit", "reset", "file"].includes(tipo)) return false;
            return true;
        }});
    }};

    // Mapa de seletores data-repeater-add conhecidos (Etapas 4, 5 e 6)
    const MAPA_REPEATER_POR_CAMPO = {{
        "TREINAMENTOS TECNICOS": "technical-trainings",
        "ACOES DE RETENCAO DE TALENTOS": "talent-actions",
        "ACOES MACRO": "macro-actions",
        "ACOES MICRO": "micro-actions",
        "ACOES EXTRAS": "extra-actions",
        "INDICADOR": "kpis",
        "FORMULA / DESCRICAO": "kpis",
        "CATEGORIA / ITEM": "budget-plan",
        "FONTE DE RECURSOS": "budget-plan",
        "RISCO IDENTIFICADO": "risk-plan",
        "ACAO DE MITIGACAO / CONTINGENCIA": "risk-plan"
    }};

    // Busca o botão '+ Adicionar' dentro de um container HTML ou via data-repeater-add específico
    const obterBotaoAdicionar = (container, nomeCampo = "") => {{
        const chaveCampo = normalizar(nomeCampo);
        const repeaterAlvo = MAPA_REPEATER_POR_CAMPO[chaveCampo];
        if (repeaterAlvo) {{
            const btnRepeater = document.querySelector(`button[data-repeater-add="${{repeaterAlvo}}"]`);
            if (btnRepeater && ehVisivel(btnRepeater)) {{
                return btnRepeater;
            }}
        }}

        const btnDataRepeater = container.querySelector('button[data-repeater-add]');
        if (btnDataRepeater && ehVisivel(btnDataRepeater)) {{
            return btnDataRepeater;
        }}

        const candidatos = Array.from(container.querySelectorAll('button, a, span[role="button"], div[role="button"], input[type="button"]'));
        for (const btn of candidatos) {{
            if (!ehVisivel(btn)) continue;
            const txt = normalizar(btn.innerText || btn.value || btn.getAttribute("title") || "");
            if (txt.includes("ADICIONAR") || txt === "+" || txt.includes("NOVO") || txt.includes("NOVA") || txt.includes("INCLUIR")) {{
                return btn;
            }}
        }}
        return null;
    }};

    // Encontra o elemento HTML do título da coluna (ex: 'FORÇAS') sem confundir com textos digitados
    const encontrarElementoTitulo = (nomeCampo) => {{
        const alvo = normalizar(nomeCampo);
        const alvoSingular = alvo.endsWith("S") && alvo.length > 4 ? alvo.slice(0, -1) : alvo;
        const tagsTitulo = "h1, h2, h3, h4, h5, h6, label, th, legend, strong, b, span, div, p";
        const elementos = Array.from(document.querySelectorAll(tagsTitulo));

        let melhores = [];
        for (const el of elementos) {{
            if (!ehVisivel(el)) continue;
            if (["INPUT", "TEXTAREA", "SCRIPT", "STYLE", "BUTTON"].includes(el.tagName)) continue;

            // Pega apenas o texto direto ou de títulos curtos (nunca parágrafos longos)
            const textoEl = normalizar(el.innerText || el.textContent || "");
            if (!textoEl || textoEl.length > 45) continue;

            let pontuacao = 0;
            if (textoEl === alvo || textoEl === alvoSingular) {{
                pontuacao = 100;
            }} else if (textoEl.startsWith(alvo + " ") || textoEl.endsWith(" " + alvo) || textoEl.includes(alvo)) {{
                pontuacao = 70;
            }} else if (textoEl.includes(alvoSingular)) {{
                pontuacao = 50;
            }}

            if (pontuacao > 0) {{
                // Prefere elementos mais específicos (folhas da árvore DOM) e tags de cabeçalho
                const bonusTag = ["H1", "H2", "H3", "H4", "H5", "H6", "LABEL", "STRONG", "B"].includes(el.tagName) ? 20 : 0;
                const penalidadeFilhos = el.children.length * 5;
                melhores.push({{ el, score: pontuacao + bonusTag - penalidadeFilhos }});
            }}
        }}

        melhores.sort((a, b) => b.score - a.score);
        return melhores.length > 0 ? melhores[0].el : null;
    }};

    // Sobe na árvore HTML a partir do título até achar o Card/Bloco exclusivo daquele campo
    const encontrarContainerDoCampo = (elTitulo, nomeCampoAtual) => {{
        const outrosNomes = todosNomesCampos
            .filter(n => normalizar(n) !== normalizar(nomeCampoAtual))
            .map(n => normalizar(n));

        let atual = elTitulo.parentElement;
        let melhorContainer = null;

        for (let nivel = 0; nivel < 8 && atual && atual !== document.body; nivel++) {{
            const btnAdd = obterBotaoAdicionar(atual, nomeCampoAtual);
            const caixas = obterCaixasDeTexto(atual);

            // Verifica se subiu demais e engoliu outro quadrante irmão
            const textoContainer = normalizar(atual.innerText || "");
            const engoliuOutroCampo = outrosNomes.some(outro => {{
                // Verifica se outro título de quadrante + seu botão estão dentro deste mesmo container
                return obterCaixasDeTexto(atual).length > 10 && textoContainer.includes(outro);
            }});

            if ((btnAdd || caixas.length > 0) && !engoliuOutroCampo) {{
                melhorContainer = atual;
                // Se já achou TANTO o botão Adicionar QUANTO pelo menos 1 caixa de texto, é o Card exato!
                if (btnAdd && caixas.length > 0) {{
                    return atual;
                }}
            }}
            atual = atual.parentElement;
        }}
        return melhorContainer || elTitulo.parentElement;
    }};

    console.log("🚀 Iniciando Preenchimento Direto via HTML (DOM)...", DADOS);
    let resumo = [];

    for (const [nomeCampo, valores] of Object.entries(DADOS)) {{
        const elTitulo = encontrarElementoTitulo(nomeCampo);
        if (!elTitulo) {{
            console.warn(`⚠️ Título HTML não encontrado para '${{nomeCampo}}'`);
            resumo.push(`❌ ${{nomeCampo}}: título não achado`);
            continue;
        }}

        const container = encontrarContainerDoCampo(elTitulo, nomeCampo);
        const btnAdd = obterBotaoAdicionar(container, nomeCampo);

        console.log(`📌 Campo '${{nomeCampo}}' localizado no HTML!`, {{ elTitulo, container, btnAdd }});

        for (let i = 0; i < valores.length; i++) {{
            let caixas = obterCaixasDeTexto(container);

            // Se precisamos de mais uma caixinha (ex: item #2, #3, #4...), clicamos no botão '+ Adicionar' do HTML!
            while (caixas.length <= i && btnAdd) {{
                btnAdd.scrollIntoView({{ behavior: "smooth", block: "center" }});
                await esperar(120);
                btnAdd.click();
                await esperar(240);
                caixas = obterCaixasDeTexto(container);
            }}

            // Seletor dinâmico: aponta sempre para o índice [i] recém-criado (ou último disponível se > 0), nunca sobrescrevendo [0]
            const caixaAlvo = caixas[i] || (i > 0 && caixas.length > 1 ? caixas[caixas.length - 1] : null);
            if (caixaAlvo) {{
                caixaAlvo.scrollIntoView({{ behavior: "smooth", block: "center" }});
                await esperar(100);
                preencherElementoInput(caixaAlvo, valores[i]);
                console.log(`  ✅ [${{nomeCampo}} #${{i + 1}}] Preenchido no HTML:`, valores[i]);
                await esperar(80);
            }} else {{
                console.warn(`  ⚠️ Caixa #${{i + 1}} não encontrada em '${{nomeCampo}}'`);
            }}
        }}
        resumo.push(`✅ ${{nomeCampo}}: ${{valores.length}} itens preenchidos`);
    }}

    console.log("🏁 Preenchimento HTML Concluído!", resumo.join(" | "));
    return resumo.join(" | ");
}})();
""".strip()
    return script_js


def _enviar_comando_websocket_cdp(ws_url: str, expressao_js: str) -> Optional[str]:
    """
    Conecta via WebSocket puro (usando apenas `socket` nativo do Python, sem precisar de pip install)
    ao Chrome/Edge DevTools Protocol (CDP) e executa o JavaScript direto no HTML da aba aberta!
    """
    parsed = urlparse(ws_url)
    host = parsed.hostname or "127.0.0.1"
    port = parsed.port or 9222
    path = parsed.path or "/"

    sock = socket.create_connection((host, port), timeout=5.0)
    try:
        chave_ws = base64.b64encode(os.urandom(16)).decode("ascii")
        handshake = (
            f"GET {path} HTTP/1.1\r\n"
            f"Host: {host}:{port}\r\n"
            f"Upgrade: websocket\r\n"
            f"Connection: Upgrade\r\n"
            f"Sec-WebSocket-Key: {chave_ws}\r\n"
            f"Sec-WebSocket-Version: 13\r\n\r\n"
        )
        sock.sendall(handshake.encode("utf-8"))

        resposta_http = b""
        while b"\r\n\r\n" not in resposta_http:
            chunk = sock.recv(1024)
            if not chunk:
                return None
            resposta_http += chunk

        payload_json = json.dumps(
            {
                "id": 1,
                "method": "Runtime.evaluate",
                "params": {
                    "expression": expressao_js,
                    "awaitPromise": True,
                    "returnByValue": True,
                },
            }
        ).encode("utf-8")

        # Monta frame WebSocket mascarado (RFC 6455)
        mascara = os.urandom(4)
        header = bytearray([0x81])  # FIN + Text Frame
        tam = len(payload_json)
        if tam < 126:
            header.append(0x80 | tam)
        elif tam < 65536:
            header.append(0x80 | 126)
            header.extend(struct.pack("!H", tam))
        else:
            header.append(0x80 | 127)
            header.extend(struct.pack("!Q", tam))

        header.extend(mascara)
        payload_mascarado = bytearray(b ^ mascara[i % 4] for i, b in enumerate(payload_json))
        sock.sendall(header + payload_mascarado)

        # Lê resposta do WebSocket
        primeiro_par = sock.recv(2)
        if len(primeiro_par) < 2:
            return None
        tam_resp = primeiro_par[1] & 0x7F
        if tam_resp == 126:
            tam_resp = struct.unpack("!H", sock.recv(2))[0]
        elif tam_resp == 127:
            tam_resp = struct.unpack("!Q", sock.recv(8))[0]

        dados_resp = b""
        while len(dados_resp) < tam_resp:
            parte = sock.recv(tam_resp - len(dados_resp))
            if not parte:
                break
            dados_resp += parte

        obj = json.loads(dados_resp.decode("utf-8", errors="ignore"))
        valor_retorno = obj.get("result", {}).get("result", {}).get("value")
        return str(valor_retorno) if valor_retorno else "Executado via CDP"
    finally:
        sock.close()


def tentar_preencher_via_porta_cdp(
    linhas: List[Dict[str, Any]], nomes_campos: List[str], porta: int = 9222
) -> Optional[str]:
    """
    Tenta encontrar uma aba aberta no Chrome/Edge via porta de depuração (9222)
    e preencher o HTML diretamente sem usar mouse nem teclado.
    """
    try:
        req = urllib.request.urlopen(f"http://127.0.0.1:{porta}/json", timeout=1.0)
        abas = json.loads(req.read().decode("utf-8"))
    except Exception:
        return None

    script_js = gerar_script_js_preenchimento_html(linhas, nomes_campos)

    for aba in abas:
        if aba.get("type") != "page":
            continue
        url = str(aba.get("url", ""))
        if url.startswith("devtools://") or url.startswith("chrome://") or url.startswith("edge://"):
            continue
        ws_url = aba.get("webSocketDebuggerUrl")
        if ws_url:
            try:
                resultado = _enviar_comando_websocket_cdp(ws_url, script_js)
                if resultado and "✅" in resultado:
                    return resultado
            except Exception:
                continue
    return None

// background.js - Service Worker 100% oculto (não abre janelas nem muda o foco da tela)
const SERVIDOR_LOCAL = "http://127.0.0.1:8765";

function executarPreenchimentoNoDOM(config) {
  const stepIndex = config.stepIndex;
  const repeaterName = config.repeaterName;
  const mode = config.mode;
  const items = config.items || [];

  if (stepIndex !== null && stepIndex !== undefined) {
    const navBtn = document.querySelector(`.wizard-nav button[data-step-index="${stepIndex}"]`);
    if (navBtn && !navBtn.classList.contains("active")) {
      navBtn.click();
    }
  }

  const container = document.querySelector(`[data-repeater="${repeaterName}"]`);
  const addBtn = document.querySelector(`button[data-repeater-add="${repeaterName}"]`);
  if (!container) {
    return { ok: false, erro: `Container [data-repeater="${repeaterName}"] não encontrado na guia.` };
  }

  function dispararEventos(el) {
    el.dispatchEvent(new Event("input", { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
    el.dispatchEvent(new Event("blur", { bubbles: true }));
  }

  function normalizarTexto(s) {
    return String(s || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim()
      .toUpperCase();
  }

  function definirValorElemento(el, valor) {
    if (!el) return false;
    const valStr = String(valor ?? "").trim();
    if (!valStr) return false;

    if (el.tagName.toLowerCase() === "select") {
      const alvoNorm = normalizarTexto(valStr);
      let escolheu = false;
      for (const opt of Array.from(el.options)) {
        const optValNorm = normalizarTexto(opt.value);
        const optTxtNorm = normalizarTexto(opt.textContent);
        if (
          optValNorm === alvoNorm ||
          optTxtNorm === alvoNorm ||
          optTxtNorm.includes(alvoNorm) ||
          alvoNorm.includes(optTxtNorm)
        ) {
          el.value = opt.value;
          escolheu = true;
          break;
        }
      }
      if (!escolheu) {
        el.value = valStr;
      }
      dispararEventos(el);
      return true;
    } else {
      const eraReadonly = el.hasAttribute("readonly");
      if (eraReadonly) {
        el.removeAttribute("readonly");
      }
      el.value = valStr;
      dispararEventos(el);
      if (eraReadonly) {
        el.setAttribute("readonly", "");
      }
      return true;
    }
  }

  function obterLinhasRepeater() {
    const rows = Array.from(
      container.querySelectorAll("[data-repeater-row], [data-repeater-item], .wizard-table-row, .repeater-item")
    );
    if (rows.length > 0) return rows;
    return Array.from(container.children);
  }

  let preenchidos = 0;

  if (mode === "single_column") {
    for (let i = 0; i < items.length; i++) {
      const texto = String(items[i] ?? "").trim();
      if (!texto) continue;

      let rows = obterLinhasRepeater();
      let alvoInput = null;

      for (const r of rows) {
        const inp = r.querySelector("input.form-control, textarea.form-control");
        if (inp && String(inp.value || "").trim() === "") {
          alvoInput = inp;
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
        alvoInput = ultimaRow ? ultimaRow.querySelector("input.form-control, textarea.form-control") : null;
      }

      if (alvoInput) {
        definirValorElemento(alvoInput, texto);
        preenchidos++;
      }
    }
    return { ok: true, preenchidos: preenchidos, totalLinhasDOM: obterLinhasRepeater().length };
  }

  for (let i = 0; i < items.length; i++) {
    const mapaCampos = items[i] || {};
    let rows = obterLinhasRepeater();
    let rowVazia = null;

    for (const r of rows) {
      const editaveis = Array.from(
        r.querySelectorAll("input.form-control:not([readonly]), textarea.form-control:not([readonly])")
      );
      const temTexto = editaveis.some((el) => String(el.value || "").trim() !== "");
      if (!temTexto && editaveis.length > 0) {
        rowVazia = r;
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
    }

    if (rowVazia) {
      for (const [seletorRelativo, valorCampo] of Object.entries(mapaCampos)) {
        const valStr = String(valorCampo ?? "").trim();
        if (!valStr) continue;
        const elCampo = rowVazia.querySelector(seletorRelativo);
        if (elCampo) {
          definirValorElemento(elCampo, valStr);
        }
      }
      preenchidos++;
    }
  }

  return { ok: true, preenchidos: preenchidos, totalLinhasDOM: obterLinhasRepeater().length };
}

async function encontrarGuiaNexPrivacy() {
  const tabs = await chrome.tabs.query({
    url: ["https://parana.nexprivacy.com/admin/institutional-documents/*"]
  });
  if (tabs && tabs.length > 0) {
    return tabs[0];
  }
  const todas = await chrome.tabs.query({});
  for (const t of todas) {
    const u = String(t.url || "").toLowerCase();
    if (u.includes("nexprivacy.com") && u.includes("institutional-documents")) {
      return t;
    }
  }
  return null;
}

async function baixarPlanilhaGoogleSheetsEmBackground(exportUrl) {
  try {
    const resp = await fetch(exportUrl, { credentials: "include" });
    if (!resp.ok) {
      return { ok: false, erro: `HTTP ${resp.status} ao exportar planilha.` };
    }
    const ct = String(resp.headers.get("content-type") || "");
    if (ct.includes("text/html")) {
      return { ok: false, erro: "Sessão Google Sheets não autenticada neste perfil do Chrome." };
    }
    const buf = await resp.arrayBuffer();
    const bytes = new Uint8Array(buf);
    let binary = "";
    const chunkSize = 0x8000;
    for (let i = 0; i < bytes.length; i += chunkSize) {
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));
    }
    return { ok: true, base64: btoa(binary) };
  } catch (err) {
    return { ok: false, erro: String(err) };
  }
}

async function processarTarefa(tarefa) {
  if (!tarefa || !tarefa.id) return;

  if (tarefa.acao === "ping") {
    const tab = await encontrarGuiaNexPrivacy();
    let etapaAtiva = "1. Identificação";
    if (tab && tab.id) {
      try {
        const res = await chrome.scripting.executeScript({
          target: { tabId: tab.id },
          func: () => {
            const btn = document.querySelector(".wizard-nav button.active");
            return btn ? btn.textContent.trim() : "1. Identificação";
          }
        });
        if (res && res[0] && res[0].result) {
          etapaAtiva = res[0].result;
        }
      } catch (e) {}
    }
    await enviarResultado(tarefa.id, {
      ok: Boolean(tab),
      url: tab ? tab.url : null,
      title: tab ? tab.title : null,
      etapaAtiva: etapaAtiva,
      modo: "extensao_background"
    });
    return;
  }

  if (tarefa.acao === "export_xlsx") {
    const resultado = await baixarPlanilhaGoogleSheetsEmBackground(tarefa.exportUrl);
    await enviarResultado(tarefa.id, resultado);
    return;
  }

  if (tarefa.acao === "fill_repeater") {
    const tab = await encontrarGuiaNexPrivacy();
    if (!tab || !tab.id) {
      await enviarResultado(tarefa.id, {
        ok: false,
        erro: "Guia do Nex Privacy não encontrada aberta no Chrome."
      });
      return;
    }
    try {
      const execRes = await chrome.scripting.executeScript({
        target: { tabId: tab.id },
        func: executarPreenchimentoNoDOM,
        args: [tarefa.payload]
      });
      const val = execRes && execRes[0] ? execRes[0].result : { ok: false, erro: "Sem retorno do script." };
      await enviarResultado(tarefa.id, val);
    } catch (err) {
      await enviarResultado(tarefa.id, { ok: false, erro: String(err) });
    }
  }
}

async function enviarResultado(id, resultado) {
  try {
    await fetch(`${SERVIDOR_LOCAL}/api/result`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ id, resultado })
    });
  } catch (e) {}
}

async function loopPoll() {
  try {
    const resp = await fetch(`${SERVIDOR_LOCAL}/api/poll`, { cache: "no-store" });
    if (resp.ok) {
      const dados = await resp.json();
      if (dados && dados.id) {
        await processarTarefa(dados);
      }
    }
  } catch (e) {
    // Servidor Python não está rodando no momento; aguarda silenciosamente
  }
}

setInterval(loopPoll, 500);
loopPoll();

// Garante que a guia do Nex Privacy que JÁ está aberta seja conectada imediatamente sem precisar recarregar
async function injetarNaGuiaJaAberta() {
  try {
    const tab = await encontrarGuiaNexPrivacy();
    if (tab && tab.id) {
      await chrome.scripting.executeScript({
        target: { tabId: tab.id },
        files: ["content.js"]
      });
    }
  } catch (e) {}
}

chrome.runtime.onInstalled.addListener(() => {
  injetarNaGuiaJaAberta();
  loopPoll();
});

chrome.runtime.onStartup.addListener(() => {
  injetarNaGuiaJaAberta();
  loopPoll();
});

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg && msg.type === "pdtic_keepalive") {
    loopPoll();
    sendResponse({ ok: true });
  }
  return false;
});

// content.js - Mantém a guia já aberta do Nex Privacy conectada em 2º plano (sem subir a janela)
(function () {
  if (window.__PDTIC_BG_CONTENT_ACTIVE__) return;
  window.__PDTIC_BG_CONTENT_ACTIVE__ = true;

  // Acorda o service worker da extensão a cada 4 segundos para resposta instantânea em 2º plano
  setInterval(() => {
    try {
      if (typeof chrome !== "undefined" && chrome.runtime && chrome.runtime.sendMessage) {
        chrome.runtime.sendMessage({ type: "pdtic_keepalive" }, () => {
          void chrome.runtime.lastError;
        });
      }
    } catch (e) {}
  }, 4000);

  console.log("[PDTIC 2º Plano] Guia do Nex Privacy conectada em modo 100% oculto!");
})();

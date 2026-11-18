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
import warnings

# Silencia os avisos visuais inofensivos do openpyxl sobre cabeçalhos e rodapés do Google Sheets
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

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
TIMEOUT_COMANDO_PADRAO = 90
DELAY_PASSO_MS_PADRAO = 450
DELAY_ADICIONAR_MS_PADRAO = 90
DELAY_CAMPO_MS_PADRAO = 25
TAMANHO_LOTE_PADRAO = 15
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

JS_HELPER_INSPECIONAR_DOM = r"""
(function(config) {
    const cfg = config || {};
    const listaRepeaters = cfg.repeaters || [
        { step: 2, name: 'strengths', label: '3.1 Forças' },
        { step: 2, name: 'weaknesses', label: '3.1 Fraquezas' },
        { step: 2, name: 'opportunities', label: '3.1 Oportunidades' },
        { step: 2, name: 'threats', label: '3.1 Ameaças' },
        { step: 2, name: 'needs', label: '3.2 Inventário de necessidades' },
        { step: 3, name: 'goals-actions', label: '4.1 Metas e ações' },
        { step: 3, name: 'budget-plan', label: '4.2 Plano orçamentário' },
        { step: 3, name: 'risk-plan', label: '4.3 Gestão de riscos' },
        { step: 4, name: 'technical-trainings', label: '5.1 Treinamentos técnicos' },
        { step: 4, name: 'talent-actions', label: '5.1 Retenção de talentos' },
        { step: 4, name: 'macro-actions', label: '5.3 Ações Macro' },
        { step: 4, name: 'micro-actions', label: '5.3 Ações Micro' },
        { step: 4, name: 'extra-actions', label: '5.3 Ações Extras' },
        { step: 5, name: 'kpis', label: '6.1 Indicadores (KPIs)' }
    ];

    function obterLinhasDiretas(container) {
        if (!container) return [];
        const candidatos = Array.from(
            container.querySelectorAll('[data-repeater-row], [data-repeater-item], .wizard-table-row, .repeater-item')
        ).filter(el => el.tagName !== 'TEMPLATE' && el.tagName !== 'SCRIPT');
        if (candidatos.length > 0) {
            return candidatos.filter(r => !candidatos.some(outro => outro !== r && outro.contains(r)));
        }
        return Array.from(container.children).filter(el => el.tagName !== 'TEMPLATE' && el.tagName !== 'SCRIPT');
    }

    function acharLabelCampo(el, row) {
        if (!el) return '';
        if (el.id) {
            const lbl = (row || document).querySelector(`label[for="${CSS.escape(el.id)}"]`);
            if (lbl && lbl.textContent.trim()) return lbl.textContent.trim();
        }
        let p = el.parentElement;
        for (let nivel = 0; nivel < 3 && p && p !== row; nivel++) {
            const lbl = p.querySelector('label, .form-label');
            if (lbl && lbl.textContent.trim()) return lbl.textContent.trim();
            p = p.parentElement;
        }
        return '';
    }

    const form = document.querySelector('#pdtic-form') || document.querySelector('form');
    let formInfo = { existe: Boolean(form) };
    if (form) {
        const camposComName = Array.from(form.querySelectorAll('input[name], select[name], textarea[name]'))
            .filter(el => !el.disabled);
        const janela1000 = [];
        for (let idx = 995; idx <= 1004 && idx < camposComName.length; idx++) {
            if (idx < 0) continue;
            const c = camposComName[idx];
            janela1000.push({
                pos1Based: idx + 1,
                tag: c.tagName.toLowerCase(),
                type: (c.getAttribute('type') || '').toLowerCase(),
                name: c.name,
                valorPreview: String(c.value || '').slice(0, 45)
            });
        }
        formInfo = {
            existe: true,
            id: form.id || '',
            method: (form.getAttribute('method') || 'GET').toUpperCase(),
            action: form.getAttribute('action') || location.href,
            enctype: form.getAttribute('enctype') || 'application/x-www-form-urlencoded',
            totalNamedFields: camposComName.length,
            janelaLimite1000: janela1000
        };
    }

    const repeatersInfo = {};
    for (const rep of listaRepeaters) {
        const container = document.querySelector(`[data-repeater="${rep.name}"]`);
        const addBtn = document.querySelector(`button[data-repeater-add="${rep.name}"]`);
        if (!container) {
            repeatersInfo[rep.name] = {
                label: rep.label,
                step: rep.step,
                exists: false,
                addBtnExists: Boolean(addBtn)
            };
            continue;
        }

        const rows = obterLinhasDiretas(container);
        let filledRows = 0;
        let emptyRows = 0;
        const partialRows = [];
        const vistosName = new Map();
        const duplicados = [];
        const amostrasPreenchidas = [];

        rows.forEach((r, rIdx) => {
            const controles = Array.from(r.querySelectorAll('input, select, textarea'))
                .filter(el => (el.getAttribute('type') || '').toLowerCase() !== 'hidden');
            controles.forEach(el => {
                if (el.name) {
                    if (vistosName.has(el.name)) {
                        duplicados.push({ name: el.name, cards: [vistosName.get(el.name), rIdx + 1] });
                    } else {
                        vistosName.set(el.name, rIdx + 1);
                    }
                }
            });

            const editaveis = controles.filter(el => !el.disabled);
            const comValor = editaveis.filter(el => String(el.value || '').trim() !== '');
            const semValor = editaveis.filter(el => String(el.value || '').trim() === '');

            if (comValor.length === 0) {
                emptyRows++;
            } else {
                filledRows++;
                const resumo = editaveis.map(el => ({
                    name: el.name || '',
                    label: acharLabelCampo(el, r),
                    valor: String(el.value || '').trim().slice(0, 40)
                }));
                amostrasPreenchidas.push({ card: rIdx + 1, qtdPreenchidos: comValor.length, totalCampos: editaveis.length, campos: resumo });
                if (semValor.length > 0 && editaveis.length > 1) {
                    partialRows.push({
                        card: rIdx + 1,
                        qtdPreenchidos: comValor.length,
                        totalCampos: editaveis.length,
                        preenchidos: comValor.map(el => acharLabelCampo(el, r) || el.name),
                        vazios: semValor.map(el => acharLabelCampo(el, r) || el.name)
                    });
                }
            }
        });

        // Schema a partir da 1ª row existente ou <template>
        let fonteSchema = rows[0] || null;
        if (!fonteSchema) {
            const tpl = container.querySelector('template') || document.querySelector(`template[data-repeater-template="${rep.name}"]`);
            if (tpl && tpl.content) fonteSchema = tpl.content;
        }
        const schema = fonteSchema
            ? Array.from(fonteSchema.querySelectorAll('input, select, textarea'))
                .filter(el => (el.getAttribute('type') || '').toLowerCase() !== 'hidden')
                .map((el, idx) => ({
                    idx: idx,
                    tag: el.tagName.toLowerCase(),
                    type: (el.getAttribute('type') || 'text').toLowerCase(),
                    name: el.getAttribute('name') || '',
                    placeholder: el.getAttribute('placeholder') || '',
                    readonly: el.hasAttribute('readonly'),
                    disabled: el.hasAttribute('disabled'),
                    label: acharLabelCampo(el, rows[0] || container)
                }))
            : [];

        repeatersInfo[rep.name] = {
            label: rep.label,
            step: rep.step,
            exists: true,
            addBtnExists: Boolean(addBtn),
            addBtnDisabled: addBtn ? Boolean(addBtn.disabled) : null,
            addBtnText: addBtn ? addBtn.textContent.trim() : '',
            totalRows: rows.length,
            filledRows: filledRows,
            emptyRows: emptyRows,
            partialRows: partialRows.slice(0, 10),
            qtdPartialRows: partialRows.length,
            duplicateNamesCount: duplicados.length,
            duplicateNamesSample: duplicados.slice(0, 5),
            schema: schema,
            ultimosPreenchidos: amostrasPreenchidas.slice(-3)
        };
    }

    const btnAtivo = document.querySelector('.wizard-nav button.active');
    return {
        ok: true,
        url: location.href,
        title: document.title,
        etapaAtiva: btnAtivo ? btnAtivo.textContent.trim() : '',
        formInfo: formInfo,
        repeaters: repeatersInfo
    };
})
"""

JS_HELPER_PREENCHER_REPEATER = r"""
(async function(config) {
    const stepIndex = config.stepIndex;
    const repeaterName = config.repeaterName;
    const mode = config.mode;
    const items = config.items || [];
    const startIndex = typeof config.startIndex === 'number' ? config.startIndex : null;
    const clearBeforeFill = Boolean(config.clearBeforeFill);
    const delayStepMs = typeof config.delayStepMs === 'number' ? config.delayStepMs : 450;
    const delayAddMs = typeof config.delayAddMs === 'number' ? config.delayAddMs : 90;
    const delayFieldMs = typeof config.delayFieldMs === 'number' ? config.delayFieldMs : 25;

    const sleep = (ms) => new Promise(resolve => setTimeout(resolve, Math.max(0, ms)));

    if (stepIndex !== null && stepIndex !== undefined) {
        const navBtn = document.querySelector(`.wizard-nav button[data-step-index="${stepIndex}"]`);
        if (navBtn && !navBtn.classList.contains('active')) {
            navBtn.click();
            await sleep(delayStepMs);
        }
    }

    let container = document.querySelector(`[data-repeater="${repeaterName}"]`);
    for (let tent = 0; !container && tent < 15; tent++) {
        await sleep(100);
        container = document.querySelector(`[data-repeater="${repeaterName}"]`);
    }
    const addBtn = document.querySelector(`button[data-repeater-add="${repeaterName}"]`);
    if (!container) {
        return {
            ok: false,
            erro: `Container HTML [data-repeater="${repeaterName}"] não foi encontrado na guia aberta (${location.href}).`
        };
    }

    function obterLinhasRepeater() {
        const candidatos = Array.from(
            container.querySelectorAll('[data-repeater-row], [data-repeater-item], .wizard-table-row, .repeater-item')
        ).filter(el => el.tagName !== 'TEMPLATE' && el.tagName !== 'SCRIPT');
        if (candidatos.length > 0) {
            return candidatos.filter(r => !candidatos.some(outro => outro !== r && outro.contains(r)));
        }
        return Array.from(container.children).filter(el => el.tagName !== 'TEMPLATE' && el.tagName !== 'SCRIPT');
    }

    function acharLabelCampo(el, row) {
        if (!el) return '';
        if (el.id) {
            const lbl = (row || document).querySelector(`label[for="${CSS.escape(el.id)}"]`);
            if (lbl && lbl.textContent.trim()) return lbl.textContent.trim();
        }
        let p = el.parentElement;
        for (let nivel = 0; nivel < 3 && p && p !== row; nivel++) {
            const lbl = p.querySelector('label, .form-label');
            if (lbl && lbl.textContent.trim()) return lbl.textContent.trim();
            p = p.parentElement;
        }
        return '';
    }

    function normalizarTexto(s) {
        return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();
    }

    function converterParaNumeroHTML(valStr) {
        let s = String(valStr || '').trim().replace(/^R\$\s*/i, '').replace(/\s+/g, '');
        if (/^-?\d{1,3}(\.\d{3})+,\d+$/.test(s)) {
            return s.replace(/\./g, '').replace(',', '.');
        }
        if (/^-?\d+,\d+$/.test(s)) {
            return s.replace(',', '.');
        }
        if (/^-?\d{1,3}(\.\d{3})+$/.test(s)) {
            return s.replace(/\./g, '');
        }
        return s;
    }

    function aplicarValorNativo(el, val) {
        const proto = el.tagName.toLowerCase() === 'textarea'
            ? window.HTMLTextAreaElement.prototype
            : window.HTMLInputElement.prototype;
        const setter = proto ? Object.getOwnPropertyDescriptor(proto, 'value')?.set : null;
        if (setter) {
            setter.call(el, val);
        } else {
            el.value = val;
        }
    }

    function dispararEventos(el) {
        el.dispatchEvent(new Event('focus', { bubbles: true }));
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true }));
        el.dispatchEvent(new Event('blur', { bubbles: true }));
    }

    async function definirValorElementoVerificado(el, valor) {
        if (!el) return { ok: false, motivo: 'elemento_nulo', valorFinal: '' };
        const valStr = String(valor ?? '').trim();
        if (!valStr) return { ok: false, motivo: 'valor_vazio', valorFinal: '' };

        const tag = el.tagName.toLowerCase();
        const tipo = (el.getAttribute('type') || 'text').toLowerCase();

        if (tag === 'select') {
            const alvoNorm = normalizarTexto(valStr);
            let escolheu = false;
            for (const opt of Array.from(el.options)) {
                const optValNorm = normalizarTexto(opt.value);
                const optTxtNorm = normalizarTexto(opt.textContent);
                if (optValNorm === alvoNorm || optTxtNorm === alvoNorm || (optTxtNorm && (optTxtNorm.includes(alvoNorm) || alvoNorm.includes(optTxtNorm)))) {
                    el.value = opt.value;
                    escolheu = true;
                    break;
                }
            }
            if (!escolheu) el.value = valStr;
            dispararEventos(el);
            if (delayFieldMs > 0) await sleep(delayFieldMs);
            const finalSel = String(el.value || '').trim();
            return { ok: finalSel !== '', valorFinal: finalSel, motivo: finalSel !== '' ? 'ok' : 'select_sem_opcao' };
        }

        const eraReadonly = el.hasAttribute('readonly');
        if (eraReadonly) el.removeAttribute('readonly');

        // Monta candidatos de formato (especialmente para inputs numéricos ou máscaras monetárias)
        const numHTML = converterParaNumeroHTML(valStr);
        const candidatos = tipo === 'number'
            ? [numHTML, valStr]
            : [valStr, numHTML, numHTML.replace('.', ',')];

        let valorLido = '';
        for (const cand of candidatos) {
            if (!cand) continue;
            aplicarValorNativo(el, cand);
            dispararEventos(el);
            if (delayFieldMs > 0) await sleep(delayFieldMs);
            valorLido = String(el.value || '').trim();
            if (valorLido !== '') break;
        }

        if (eraReadonly) el.setAttribute('readonly', '');
        return {
            ok: valorLido !== '',
            valorFinal: valorLido,
            motivo: valorLido !== '' ? 'ok' : `rejeitado_pelo_input(type=${tipo})`
        };
    }

    async function garantirCardNoIndice(targetIdx) {
        let rows = obterLinhasRepeater();
        let criadosAgora = 0;
        while (rows.length <= targetIdx) {
            if (!addBtn) {
                return {
                    ok: false,
                    criadosAgora,
                    erro: `Botão [data-repeater-add="${repeaterName}"] não encontrado ao tentar criar card #${targetIdx + 1} (DOM possui apenas ${rows.length} card(s)).`
                };
            }
            if (addBtn.disabled) {
                return {
                    ok: false,
                    criadosAgora,
                    capacidadeMaximaAtingida: true,
                    erro: `Botão [data-repeater-add="${repeaterName}"] está DESABILITADO no card #${rows.length} (impossível criar card #${targetIdx + 1}).`
                };
            }
            const qtdAntes = rows.length;
            addBtn.click();
            await sleep(delayAddMs);
            rows = obterLinhasRepeater();
            for (let espera = 0; rows.length <= qtdAntes && espera < 14; espera++) {
                await sleep(50);
                rows = obterLinhasRepeater();
            }
            if (rows.length <= qtdAntes) {
                return {
                    ok: false,
                    criadosAgora,
                    capacidadeMaximaAtingida: true,
                    erro: `Clique em '+ Adicionar' (${repeaterName}) NÃO criou novo card! Capacidade travou em ${rows.length} card(s) ao tentar criar o card #${targetIdx + 1}.`
                };
            }
            criadosAgora++;
        }
        return { ok: true, row: rows[targetIdx], criadosAgora, totalAgora: rows.length };
    }

    function corrigirIndicesDuplicadosNoRepeater() {
        const rows = obterLinhasRepeater();
        let corrigidos = 0;
        const vistos = new Set();
        let temColisao = false;
        for (const r of rows) {
            for (const el of Array.from(r.querySelectorAll('input[name], select[name], textarea[name]'))) {
                if (vistos.has(el.name)) {
                    temColisao = true;
                    break;
                }
                vistos.add(el.name);
            }
            if (temColisao) break;
        }
        if (!temColisao) return 0;

        rows.forEach((r, idxCorreto) => {
            const campos = Array.from(r.querySelectorAll('input[name], select[name], textarea[name]'));
            campos.forEach(el => {
                const atual = el.name;
                let novo = atual;
                if (/\[\d+\](\[[^\]]+\])$/.test(atual)) {
                    novo = atual.replace(/\[\d+\](\[[^\]]+\])$/, `[${idxCorreto}]$1`);
                } else if (/\[\d+\]$/.test(atual)) {
                    novo = atual.replace(/\[\d+\]$/, `[${idxCorreto}]`);
                }
                if (novo !== atual) {
                    el.name = novo;
                    corrigidos++;
                }
            });
        });
        return corrigidos;
    }

    if (clearBeforeFill && (startIndex === 0 || startIndex === null)) {
        let rows = obterLinhasRepeater();
        for (let idx = rows.length - 1; idx >= 1; idx--) {
            const rmBtn = rows[idx].querySelector('[data-repeater-remove], [data-repeater-delete], button.btn-danger, .remove-row');
            if (rmBtn) {
                rmBtn.click();
                await sleep(15);
            }
        }
        rows = obterLinhasRepeater();
        rows.forEach(r => {
            Array.from(r.querySelectorAll('input:not([type="hidden"]), textarea')).forEach(el => {
                if (!el.disabled) el.value = '';
            });
        });
        await sleep(80);
    }

    const linhasIniciaisDOM = obterLinhasRepeater().length;
    let totalCardsCriados = 0;
    let preenchidos = 0;
    let camposEsperadosTotal = 0;
    let camposConfirmadosTotal = 0;
    const detalhes = [];
    const seletoresNaoEncontrados = [];
    const camposRejeitados = [];
    const mapeamentoRealCampos = {};

    if (mode === 'single_column') {
        for (let i = 0; i < items.length; i++) {
            const texto = String(items[i] ?? '').trim();
            if (!texto) continue;
            camposEsperadosTotal++;

            let alvoInput = null;
            let idxUsado = -1;

            if (startIndex !== null) {
                const targetIdx = startIndex + i;
                const gar = await garantirCardNoIndice(targetIdx);
                totalCardsCriados += gar.criadosAgora || 0;
                if (!gar.ok) {
                    return {
                        ok: false,
                        erro: gar.erro,
                        preenchidos,
                        cardsCriados: totalCardsCriados,
                        linhasIniciaisDOM,
                        totalLinhasDOM: obterLinhasRepeater().length,
                        detalhes
                    };
                }
                idxUsado = targetIdx + 1;
                alvoInput = gar.row ? gar.row.querySelector('input.form-control, textarea.form-control, input:not([type="hidden"])') : null;
            } else {
                let rows = obterLinhasRepeater();
                for (let rIdx = 0; rIdx < rows.length; rIdx++) {
                    const inp = rows[rIdx].querySelector('input.form-control, textarea.form-control, input:not([type="hidden"])');
                    if (inp && String(inp.value || '').trim() === '') {
                        alvoInput = inp;
                        idxUsado = rIdx + 1;
                        break;
                    }
                }
                if (!alvoInput) {
                    const gar = await garantirCardNoIndice(rows.length);
                    totalCardsCriados += gar.criadosAgora || 0;
                    if (!gar.ok) {
                        return { ok: false, erro: gar.erro, preenchidos, cardsCriados: totalCardsCriados, totalLinhasDOM: obterLinhasRepeater().length, detalhes };
                    }
                    idxUsado = gar.totalAgora;
                    alvoInput = gar.row ? gar.row.querySelector('input.form-control, textarea.form-control, input:not([type="hidden"])') : null;
                }
            }

            if (alvoInput) {
                const verif = await definirValorElementoVerificado(alvoInput, texto);
                if (verif.ok) {
                    preenchidos++;
                    camposConfirmadosTotal++;
                    detalhes.push(`caixa #${idxUsado} (${alvoInput.name || 'input'}) <- "${verif.valorFinal.slice(0, 38)}"`);
                } else {
                    camposRejeitados.push({ card: idxUsado, seletor: 'single_column', motivo: verif.motivo, valorTentado: texto.slice(0, 40) });
                }
            } else {
                seletoresNaoEncontrados.push({ card: idxUsado, seletor: 'input.form-control, textarea.form-control' });
            }
        }

        const indicesCorrigidos = corrigirIndicesDuplicadosNoRepeater();
        const form = document.querySelector('#pdtic-form') || document.querySelector('form');
        const totalNamedForm = form ? form.querySelectorAll('input[name]:not(:disabled), select[name]:not(:disabled), textarea[name]:not(:disabled)').length : 0;

        return {
            ok: true,
            preenchidos,
            linhasIniciaisDOM,
            cardsCriados: totalCardsCriados,
            totalLinhasDOM: obterLinhasRepeater().length,
            camposEsperadosTotal,
            camposConfirmadosTotal,
            seletoresNaoEncontrados,
            camposRejeitados,
            indicesCorrigidos,
            totalNamedForm,
            detalhes
        };
    }

    for (let i = 0; i < items.length; i++) {
        const mapaCampos = items[i] || {};
        let rowAlvo = null;
        let idxUsado = -1;

        if (startIndex !== null) {
            const targetIdx = startIndex + i;
            const gar = await garantirCardNoIndice(targetIdx);
            totalCardsCriados += gar.criadosAgora || 0;
            if (!gar.ok) {
                return {
                    ok: false,
                    erro: gar.erro,
                    preenchidos,
                    linhasIniciaisDOM,
                    cardsCriados: totalCardsCriados,
                    totalLinhasDOM: obterLinhasRepeater().length,
                    seletoresNaoEncontrados,
                    camposRejeitados,
                    detalhes
                };
            }
            rowAlvo = gar.row;
            idxUsado = targetIdx + 1;
        } else {
            let rows = obterLinhasRepeater();
            for (let rIdx = 0; rIdx < rows.length; rIdx++) {
                const r = rows[rIdx];
                const editaveis = Array.from(r.querySelectorAll('input.form-control:not([readonly]), textarea.form-control:not([readonly])'));
                const temTexto = editaveis.some(el => String(el.value || '').trim() !== '');
                if (!temTexto && editaveis.length > 0) {
                    rowAlvo = r;
                    idxUsado = rIdx + 1;
                    break;
                }
            }
            if (!rowAlvo) {
                const gar = await garantirCardNoIndice(rows.length);
                totalCardsCriados += gar.criadosAgora || 0;
                if (!gar.ok) {
                    return { ok: false, erro: gar.erro, preenchidos, cardsCriados: totalCardsCriados, totalLinhasDOM: obterLinhasRepeater().length, detalhes };
                }
                rowAlvo = gar.row;
                idxUsado = gar.totalAgora;
            }
        }

        if (rowAlvo) {
            let camposOk = 0;
            let camposTentadosCard = 0;
            const falhasCard = [];

            for (const [seletorRelativo, valorCampo] of Object.entries(mapaCampos)) {
                const valStr = String(valorCampo ?? '').trim();
                if (!valStr) continue;
                camposTentadosCard++;
                camposEsperadosTotal++;

                const elCampo = rowAlvo.querySelector(seletorRelativo);
                if (!elCampo) {
                    const disponiveis = Array.from(rowAlvo.querySelectorAll('input:not([type="hidden"]), select, textarea'))
                        .map(e => `${e.tagName.toLowerCase()}[type=${e.type || 'text'} name="${e.name || ''}"]`);
                    seletoresNaoEncontrados.push({
                        card: idxUsado,
                        seletor: seletorRelativo,
                        camposDisponiveisNoCard: disponiveis
                    });
                    falhasCard.push(`NÃO_ACHADO(${seletorRelativo})`);
                    continue;
                }

                if (!mapeamentoRealCampos[seletorRelativo]) {
                    mapeamentoRealCampos[seletorRelativo] = `${elCampo.tagName.toLowerCase()}[type="${(elCampo.getAttribute('type') || 'text').toLowerCase()}" name="${elCampo.name || ''}" label="${acharLabelCampo(elCampo, rowAlvo)}"]`;
                }

                const verif = await definirValorElementoVerificado(elCampo, valStr);
                if (verif.ok) {
                    camposOk++;
                    camposConfirmadosTotal++;
                } else {
                    camposRejeitados.push({
                        card: idxUsado,
                        seletor: seletorRelativo,
                        name: elCampo.name || '',
                        type: elCampo.type || '',
                        motivo: verif.motivo,
                        valorTentado: valStr.slice(0, 40)
                    });
                    falhasCard.push(`REJEITADO(${elCampo.name || seletorRelativo}: ${verif.motivo})`);
                }
            }

            preenchidos++;
            const sufixoFalha = falhasCard.length > 0 ? ` ⚠️ FALHAS: ${falhasCard.join(', ')}` : '';
            detalhes.push(`sub-card #${idxUsado}: ${camposOk}/${camposTentadosCard} campo(s) confirmado(s) no DOM${sufixoFalha}`);
        }
    }

    const indicesCorrigidos = corrigirIndicesDuplicadosNoRepeater();
    const form = document.querySelector('#pdtic-form') || document.querySelector('form');
    const totalNamedForm = form ? form.querySelectorAll('input[name]:not(:disabled), select[name]:not(:disabled), textarea[name]:not(:disabled)').length : 0;

    return {
        ok: true,
        preenchidos,
        linhasIniciaisDOM,
        cardsCriados: totalCardsCriados,
        totalLinhasDOM: obterLinhasRepeater().length,
        camposEsperadosTotal,
        camposConfirmadosTotal,
        mapeamentoRealCampos,
        seletoresNaoEncontrados: seletoresNaoEncontrados.slice(0, 15),
        qtdSeletoresNaoEncontrados: seletoresNaoEncontrados.length,
        camposRejeitados: camposRejeitados.slice(0, 15),
        qtdCamposRejeitados: camposRejeitados.length,
        indicesCorrigidos,
        totalNamedForm,
        detalhes
    };
})
"""

JS_BRIDGE_CONSOLE = (
    r"""
(function() {
    if (window.__PDTIC_BRIDGE_TIMER_ID__) {
        clearInterval(window.__PDTIC_BRIDGE_TIMER_ID__);
        console.log("🔄 [PDTIC 2º Plano] Atualizando ponte ativa para a versão com diagnóstico avançado...");
    }
    window.__PDTIC_BRIDGE_RUNNING__ = true;
    const SERVIDOR = "__SERVIDOR_URL__";
    const fnInspecionar = """
    + JS_HELPER_INSPECIONAR_DOM.strip()
    + r""";
    const fnPreencher = """
    + JS_HELPER_PREENCHER_REPEATER.strip()
    + r""";

    let ocupada = false;

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
        if (ocupada) return;
        ocupada = true;
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
            } else if (tarefa.acao === 'inspect_dom') {
                const res = fnInspecionar(tarefa.payload || {});
                await enviarResultado(tarefa.id, res);
            } else if (tarefa.acao === 'fill_repeater') {
                const res = await fnPreencher(tarefa.payload);
                await enviarResultado(tarefa.id, res);
            }
        } catch (e) {
            console.error("[PDTIC 2º Plano] Erro no poll:", e);
        } finally {
            ocupada = false;
        }
    }

    window.__PDTIC_BRIDGE_TIMER_ID__ = setInterval(poll, 350);
    poll();
    console.log("✅ [PDTIC 2º Plano v2] Guia conectada com inspeção de campos e verificação reversa!");
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

    def __init__(
        self,
        porta_ponte: int = PORTA_PONTE_LOCAL,
        porta_cdp: int = PORTA_CDP_PADRAO,
        delay_step_ms: int = DELAY_PASSO_MS_PADRAO,
        delay_add_ms: int = DELAY_ADICIONAR_MS_PADRAO,
        delay_field_ms: int = DELAY_CAMPO_MS_PADRAO,
        tamanho_lote: int = TAMANHO_LOTE_PADRAO,
        limpar_antes: bool = False,
    ) -> None:
        self.porta_ponte = porta_ponte
        self.porta_cdp = porta_cdp
        self.delay_step_ms = delay_step_ms
        self.delay_add_ms = delay_add_ms
        self.delay_field_ms = delay_field_ms
        self.tamanho_lote = max(1, tamanho_lote)
        self.limpar_antes = limpar_antes
        self.cliente_cdp: Optional[ClienteCDP] = None
        self.modo_ativo: str = "nenhum"
        self.historico_relatorio: List[Dict[str, Any]] = []
        self.ultimo_diagnostico_dom: Optional[Dict[str, Any]] = None

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

    def gerar_relatorio_auditoria(self) -> Optional[Path]:
        """Gera um arquivo Markdown cruzando os dados lidos da planilha com os colados e confirmados no DOM."""
        if not self.historico_relatorio:
            return None

        DADOS_DIR.mkdir(parents=True, exist_ok=True)
        caminho_relatorio = DADOS_DIR / f"relatorio_pdtic_{time.strftime('%Y%m%d_%H%M%S')}.md"

        linhas_md = [
            "# 📊 Relatório de Auditoria - Automação PDTIC",
            f"**Data da Execução:** {time.strftime('%d/%m/%Y às %H:%M:%S')}",
            f"**Modo de Conexão Usado:** {self.modo_ativo.upper()}",
            f"**Tempos de Espera:** Passo={self.delay_step_ms}ms | +Adicionar={self.delay_add_ms}ms | Campo={self.delay_field_ms}ms | Lote={self.tamanho_lote}",
            "---",
            "## 🎯 Resumo Geral",
            f"- **Blocos Processados:** {len(self.historico_relatorio)}",
            f"- **Total de Linhas da Planilha:** {sum(h['total_planilha'] for h in self.historico_relatorio)}",
            f"- **Total de Itens Confirmados no DOM:** {sum(h['preenchidos'] for h in self.historico_relatorio)}",
            f"- **Novos Sub-cards Criados (+Adicionar):** {sum(h.get('cards_criados', 0) for h in self.historico_relatorio)}",
            f"- **Campos Confirmados (Leitura Reversa):** {sum(h.get('campos_confirmados', 0) for h in self.historico_relatorio)} / {sum(h.get('campos_esperados', 0) for h in self.historico_relatorio)}",
        ]
        ultimo_total_form = max((h.get("total_named_form", 0) for h in self.historico_relatorio), default=0)
        if ultimo_total_form:
            linhas_md.append(f"- **Total Final de Variáveis `[name]` no `<form>`:** {ultimo_total_form}")
            if ultimo_total_form > 1000:
                linhas_md.append(
                    f"- ⚠️ **ALERTA DE SERVIDOR (`max_input_vars`):** O `<form>` possui **{ultimo_total_form} campos `[name]`** "
                    "(acima de 1.000). Se o servidor PHP estiver com `max_input_vars = 1000`, ao clicar em Salvar o backend truncará no 1.000º campo!"
                )
        linhas_md.extend([
            "---",
            "## 📝 Detalhamento por Bloco (De -> Para)\n"
        ])

        for hist in self.historico_relatorio:
            linhas_md.append(f"### {hist['bloco']}")
            linhas_md.append(f"- **Contêiner Alvo (Repeater):** `{hist['repeater']}`")
            linhas_md.append(f"- **Comparativo Planilha × DOM:** {hist['preenchidos']} de {hist['total_planilha']} linha(s) | DOM Inicial={hist.get('dom_inicial', '?')} -> Criados=+{hist.get('cards_criados', 0)} -> DOM Final={hist.get('dom_final', '?')}")
            linhas_md.append(f"- **Campos Verificados no DOM:** {hist.get('campos_confirmados', 0)} de {hist.get('campos_esperados', 0)} campo(s)")
            linhas_md.append(f"- **Tempo Gasto:** {hist['tempo']}s")

            map_campos = hist.get("mapeamento_campos") or {}
            if map_campos:
                linhas_md.append("\n**🔎 Mapeamento HTML Real Detectado:**")
                for sel_k, desc_html in map_campos.items():
                    linhas_md.append(f"> - `{sel_k}` → `{desc_html}`")

            nao_enc = hist.get("seletores_nao_encontrados") or []
            rejeit = hist.get("campos_rejeitados") or []
            if nao_enc or rejeit:
                linhas_md.append("\n**⚠️ Alertas de Campos neste Bloco:**")
                for ne in nao_enc[:10]:
                    linhas_md.append(f"> - ❌ Seletor não encontrado no card #{ne.get('card')}: `{ne.get('seletor')}` (disponíveis: `{ne.get('camposDisponiveisNoCard')}`)")
                for rj in rejeit[:10]:
                    linhas_md.append(f"> - ⚠️ Valor rejeitado no card #{rj.get('card')} (`{rj.get('name')}` type=`{rj.get('type')}`): `{rj.get('motivo')}` | valor=`{rj.get('valorTentado')}`")
            
            # --- SEÇÃO 1: O QUE VEIO DA PLANILHA ---
            linhas_md.append("\n**📄 1. O QUE FOI LIDO DA PLANILHA:**")
            itens_enviados = hist.get('itens_enviados', [])
            
            if not itens_enviados:
                linhas_md.append("> *Nenhum dado lido.*")
            else:
                for idx, item in enumerate(itens_enviados, 1):
                    if isinstance(item, dict):
                        valores = [f"`{str(v).strip()}`" for v in item.values() if str(v).strip()]
                        texto_item = " | ".join(valores) if valores else "*Linha vazia*"
                        linhas_md.append(f"> **Linha {idx}:** {texto_item}")
                    else:
                        linhas_md.append(f"> **Linha {idx}:** `{str(item).strip()}`")

            # --- SEÇÃO 2: O QUE O JAVASCRIPT COLOCOU NA TELA ---
            linhas_md.append("\n**📋 2. O QUE FOI COLADO NO SISTEMA (Confirmação):**")
            if not hist['detalhes']:
                linhas_md.append("> *Nenhum detalhe retornado do navegador.*")
            else:
                for det in hist['detalhes']:
                    linhas_md.append(f"> - {det}")
            linhas_md.append("\n---\n")

        caminho_relatorio.write_text("\n".join(linhas_md), encoding="utf-8")
        return caminho_relatorio

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

    def inspecionar_dom(
        self,
        payload: Optional[Dict[str, Any]] = None,
        timeout: float = 25.0,
    ) -> Dict[str, Any]:
        """Inspeciona todos os repeaters, campos HTML, cards parciais e total de inputs do <form>."""
        cfg = payload or {}
        if self.modo_ativo == "cdp" and self.cliente_cdp is not None:
            try:
                expressao = f"{JS_HELPER_INSPECIONAR_DOM}({json.dumps(cfg, ensure_ascii=False)})"
                res = self.cliente_cdp.executar_js(expressao)
                if isinstance(res, dict):
                    self.ultimo_diagnostico_dom = res
                    return res
                return {"ok": False, "erro": f"Retorno inesperado do inspect_dom via CDP: {res}"}
            except Exception as exc:
                return {"ok": False, "erro": f"Exceção CDP em inspect_dom: {exc}"}

        res_ponte = self.enviar_comando_ponte("inspect_dom", {"payload": cfg}, timeout=timeout)
        if not isinstance(res_ponte, dict):
            return {
                "ok": False,
                "erro": (
                    "A ponte ativa na guia ainda é da versão anterior (ou não respondeu ao inspect_dom). "
                    "Recole o comando `fetch('http://127.0.0.1:8765/bridge.js').then(r=>r.text()).then(eval)` no Console [F12]."
                ),
            }
        self.ultimo_diagnostico_dom = res_ponte
        return res_ponte

    def _executar_chamada_repeater_unica(
        self,
        payload: Dict[str, Any],
        timeout: float = float(TIMEOUT_COMANDO_PADRAO),
    ) -> Dict[str, Any]:
        t0 = time.time()
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
                return {"ok": False, "erro": f"Exceção CDP durante preenchimento: {exc}"}

        res_ponte = self.enviar_comando_ponte("fill_repeater", {"payload": payload}, timeout=timeout)
        dt = time.time() - t0
        if not isinstance(res_ponte, dict):
            return {
                "ok": False,
                "erro": f"TIMEOUT ({timeout:.0f}s) aguardando resposta do navegador.",
            }
        res_ponte["_tempo_s"] = round(dt, 2)
        return res_ponte

    def preencher_repeater(
        self,
        payload: Dict[str, Any],
        nome_exibicao: str,
        timeout: float = float(TIMEOUT_COMANDO_PADRAO),
    ) -> Dict[str, Any]:
        """
        Preenche um repeater em lotes cadenciados (com tempo de espera entre cards e entre lotes),
        analisando previamente o DOM e comparando:
        Linhas na Planilha × Cards Iniciais no DOM × Novos Cards Criados × Campos Confirmados por Leitura Reversa!
        """
        todos_itens = list(payload.get("items") or [])
        total_planilha = len(todos_itens)
        repeater_name = str(payload.get("repeaterName") or "")
        step_index = payload.get("stepIndex")
        mode = str(payload.get("mode") or "multi_field")

        # 1. Pré-análise do Repeater no DOM antes de colar
        info_antes = self.inspecionar_dom(
            {"repeaters": [{"step": step_index, "name": repeater_name, "label": nome_exibicao}]},
            timeout=10.0,
        )
        dom_inicial_analise = "?"
        if info_antes.get("ok") and isinstance(info_antes.get("repeaters"), dict):
            rep_pre = info_antes["repeaters"].get(repeater_name) or {}
            if rep_pre.get("exists"):
                tot_r = rep_pre.get("totalRows", 0)
                fill_r = rep_pre.get("filledRows", 0)
                emp_r = rep_pre.get("emptyRows", 0)
                part_r = rep_pre.get("qtdPartialRows", 0)
                dom_inicial_analise = tot_r
                prev_novos = max(0, total_planilha - tot_r)
                btn_status = (
                    "DESABILITADO ❌"
                    if rep_pre.get("addBtnDisabled")
                    else ("OK ✅" if rep_pre.get("addBtnExists") else "NÃO ENCONTRADO ❌")
                )
                log_debug(
                    nome_exibicao,
                    f"📊 ANÁLISE PRÉ-COLAGEM ({repeater_name}): "
                    f"Linhas na Planilha={total_planilha} | Cards já no DOM={tot_r} "
                    f"(preenchidos={fill_r}, vazios={emp_r}, parciais={part_r}) | "
                    f"Novos cards a criar=+{prev_novos} | Botão '+Adicionar'={btn_status}",
                    nivel="INFO",
                )
                schema = rep_pre.get("schema") or []
                if schema:
                    resumo_schema = ", ".join(
                        f"{c.get('tag')}[type={c.get('type')} name='{c.get('name')}'{' readonly' if c.get('readonly') else ''}]"
                        for c in schema
                    )
                    log_debug(nome_exibicao, f"🔎 Campos HTML do card ({len(schema)}): {resumo_schema}", nivel="DEBUG")
                if part_r > 0:
                    log_debug(
                        nome_exibicao,
                        f"⚠️ Detectado(s) {part_r} card(s) PARCIALMENTE preenchido(s) no DOM antes de iniciar: {rep_pre.get('partialRows')}",
                        nivel="WARN",
                    )

        # 2. Execução em lotes cadenciados (evita travamento/confusão no DOM em listas grandes)
        tam_lote = self.tamanho_lote
        total_lotes = max(1, (total_planilha + tam_lote - 1) // tam_lote)
        t_global = time.time()

        preenchidos_acum = 0
        cards_criados_acum = 0
        campos_esp_acum = 0
        campos_conf_acum = 0
        detalhes_acum: List[str] = []
        mapeamento_real: Dict[str, str] = {}
        seletores_nao_enc: List[Dict[str, Any]] = []
        campos_rejeitados: List[Dict[str, Any]] = []
        dom_inicial_real: Any = dom_inicial_analise
        dom_final_real: Any = 0
        total_named_form = 0

        for idx_lote in range(total_lotes):
            ini = idx_lote * tam_lote
            fim = min(total_planilha, ini + tam_lote)
            fatia = todos_itens[ini:fim]

            payload_lote = {
                "stepIndex": step_index,
                "repeaterName": repeater_name,
                "mode": mode,
                "items": fatia,
                "startIndex": ini,
                "clearBeforeFill": bool(self.limpar_antes and idx_lote == 0),
                "delayStepMs": self.delay_step_ms if idx_lote == 0 else 80,
                "delayAddMs": self.delay_add_ms,
                "delayFieldMs": self.delay_field_ms,
            }

            res_lote = self._executar_chamada_repeater_unica(payload_lote, timeout=timeout)
            if not isinstance(res_lote, dict) or not res_lote.get("ok"):
                erro_msg = res_lote.get("erro") if isinstance(res_lote, dict) else str(res_lote)
                log_debug(
                    nome_exibicao,
                    f"❌ Falha no Lote {idx_lote + 1}/{total_lotes} (linhas {ini + 1}..{fim} de {total_planilha}): {erro_msg}",
                    nivel="ERRO",
                )
                return {
                    "ok": False,
                    "erro": erro_msg,
                    "loteFalhou": idx_lote + 1,
                    "linhasProcessadasAntesDoErro": preenchidos_acum,
                    "cardsCriadosAntesDoErro": cards_criados_acum,
                    "detalhes": detalhes_acum,
                }

            if idx_lote == 0 and "linhasIniciaisDOM" in res_lote:
                dom_inicial_real = res_lote.get("linhasIniciaisDOM")

            p_lote = int(res_lote.get("preenchidos", 0))
            c_lote = int(res_lote.get("cardsCriados", 0))
            ce_lote = int(res_lote.get("camposEsperadosTotal", 0))
            cc_lote = int(res_lote.get("camposConfirmadosTotal", 0))
            dom_final_real = res_lote.get("totalLinhasDOM", dom_final_real)
            total_named_form = int(res_lote.get("totalNamedForm", total_named_form) or 0)

            preenchidos_acum += p_lote
            cards_criados_acum += c_lote
            campos_esp_acum += ce_lote
            campos_conf_acum += cc_lote
            detalhes_acum.extend(res_lote.get("detalhes") or [])
            mapeamento_real.update(res_lote.get("mapeamentoRealCampos") or {})
            seletores_nao_enc.extend(res_lote.get("seletoresNaoEncontrados") or [])
            campos_rejeitados.extend(res_lote.get("camposRejeitados") or [])

            log_debug(
                nome_exibicao,
                f"📦 Lote {idx_lote + 1}/{total_lotes} (linhas {ini + 1}..{fim}/{total_planilha}) em {res_lote.get('_tempo_s', 0)}s | "
                f"+{c_lote} card(s) criado(s) | DOM atual={dom_final_real} | "
                f"Campos confirmados={cc_lote}/{ce_lote} | Total inputs <form>={total_named_form}",
                nivel="DEBUG" if total_lotes > 1 else "INFO",
            )

            # Pausa entre lotes para o navegador respirar e atualizar layout/máscaras
            if idx_lote + 1 < total_lotes:
                time.sleep(0.30)

        dt_total = round(time.time() - t_global, 2)

        if seletores_nao_enc:
            log_debug(
                nome_exibicao,
                f"❌ ATENÇÃO: {len(seletores_nao_enc)} seletor(es) NÃO encontrado(s) no HTML! Exemplo: {seletores_nao_enc[0]}",
                nivel="ERRO",
            )
        if campos_rejeitados:
            log_debug(
                nome_exibicao,
                f"⚠️ ATENÇÃO: {len(campos_rejeitados)} campo(s) ficaram vazios após digitação! Exemplo: {campos_rejeitados[0]}",
                nivel="WARN",
            )
        if preenchidos_acum != total_planilha:
            log_debug(
                nome_exibicao,
                f"⚠️ DIVERGÊNCIA DE QUANTIDADE: Planilha tinha {total_planilha} linha(s), mas {preenchidos_acum} foram confirmadas!",
                nivel="WARN",
            )

        self.historico_relatorio.append({
            "bloco": nome_exibicao,
            "repeater": repeater_name,
            "total_planilha": total_planilha,
            "preenchidos": preenchidos_acum,
            "dom_inicial": dom_inicial_real,
            "cards_criados": cards_criados_acum,
            "dom_final": dom_final_real,
            "campos_esperados": campos_esp_acum,
            "campos_confirmados": campos_conf_acum,
            "mapeamento_campos": mapeamento_real,
            "seletores_nao_encontrados": seletores_nao_enc,
            "campos_rejeitados": campos_rejeitados,
            "total_named_form": total_named_form,
            "detalhes": detalhes_acum,
            "tempo": dt_total,
            "itens_enviados": todos_itens,
        })

        return {
            "ok": True,
            "preenchidos": preenchidos_acum,
            "cardsCriados": cards_criados_acum,
            "totalLinhasDOM": dom_final_real,
            "camposEsperadosTotal": campos_esp_acum,
            "camposConfirmadosTotal": campos_conf_acum,
            "seletoresNaoEncontrados": seletores_nao_enc,
            "camposRejeitados": campos_rejeitados,
            "totalNamedForm": total_named_form,
            "detalhes": detalhes_acum,
            "_tempo_s": dt_total,
        }


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

    for det in (res.get("detalhes") or [])[:6]:
        log_debug(nome_exibicao, f"   ↳ {det}", nivel="DEBUG")
    log_debug(
        nome_exibicao,
        f"Concluído em {res.get('_tempo_s', 0)}s: Planilha={len(valores_lim)} | "
        f"Confirmados={res.get('preenchidos')}/{len(valores_lim)} | "
        f"Criados=+{res.get('cardsCriados', 0)} | Total DOM={res.get('totalLinhasDOM')}!",
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

    for det in (res.get("detalhes") or [])[:6]:
        log_debug(nome_exibicao, f"   ↳ {det}", nivel="DEBUG")
    log_debug(
        nome_exibicao,
        f"Concluído em {res.get('_tempo_s', 0)}s: Planilha={len(itens_lim)} | "
        f"Sub-cards confirmados={res.get('preenchidos')}/{len(itens_lim)} | "
        f"Criados=+{res.get('cardsCriados', 0)} | Total DOM={res.get('totalLinhasDOM')} | "
        f"Campos={res.get('camposConfirmadosTotal', 0)}/{res.get('camposEsperadosTotal', 0)}!",
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
            '[name$="[description]"]': buscar_valor_da_coluna(l, "DESCRIÇÃO"),
            '[name$="[priority]"]': buscar_valor_da_coluna(l, "PRIORIDADE"),
            '[name$="[requester_area]"], [name$="[requesting_area]"]': buscar_valor_da_coluna(l, "ÁREA DEMANDANTE"),
            '[name$="[justification]"]': buscar_valor_da_coluna(l, "JUSTIFICATIVA"),
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
            '[name$="[id]"]': buscar_valor_da_coluna(l, "ID"),
            '[name$="[strategic_objective]"]': buscar_valor_da_coluna(l, "OBJETIVO ESTRATÉGICO"),
            '[name$="[goal]"]': buscar_valor_da_coluna(l, "META"),
            '[name$="[main_action]"]': buscar_valor_da_coluna(l, "AÇÃO PRINCIPAL"),
            '[name$="[indicator]"]': buscar_valor_da_coluna(l, "INDICADOR"),
            '[name$="[deadline]"]': buscar_valor_da_coluna(l, "PRAZO"),
            '[name$="[responsible]"]': buscar_valor_da_coluna(l, "RESPONSÁVEL"),
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
        val_fonte = buscar_valor_da_coluna(l, "FONTE DE RECURSO")
        item = {
            '[name$="[action_id]"], [name$="[id]"]': buscar_valor_da_coluna(l, "AÇÃO (ID)"),
            '[name$="[cost_description]"], [name$="[description]"], [name$="[item]"]': buscar_valor_da_coluna(l, "DESCRIÇÃO DO CUSTO"),
            '[name$="[funding_source]"], [name$="[resource_source]"], [name*="funding"], [name*="source"], [name*="fonte"], [placeholder*="Tesouro"], [placeholder*="Convênio"]': val_fonte,
            '[name$="[year_1]"], [name$="[year1]"], [name*="2027"]': buscar_valor_da_coluna(l, "ANO 1"),
            '[name$="[year_2]"], [name$="[year2]"], [name*="2028"]': buscar_valor_da_coluna(l, "ANO 2"),
            '[name$="[total]"]': buscar_valor_da_coluna(l, "TOTAL"),
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
                    '[name$="[name]"]': ind,
                    '[name$="[formula]"]': form,
                }
            )
    return _preencher_repeater_grupo(
        ctrl, 5, "kpis", "6. Indicadores de sucesso (KPIs)", itens, limite
    )


def _contar_linhas_planilha_por_repeater(
    caminho_xlsx: Optional[Path],
    limite: Optional[int] = None,
) -> Dict[str, int]:
    """Lê rapidamente a contagem de linhas válidas por repeater na planilha para comparar com o DOM."""
    if caminho_xlsx is None or not caminho_xlsx.is_file():
        return {}
    contagens: Dict[str, int] = {}

    swot = ler_guia_de_arquivo_xlsx(caminho_xlsx, "3.1. SWOT", list(COLUNAS_PADRAO_SWOT))
    for nome_col, rep_name in [
        ("FORÇAS", "strengths"),
        ("FRAQUEZAS", "weaknesses"),
        ("OPORTUNIDADES", "opportunities"),
        ("AMEAÇAS", "threats"),
    ]:
        vals = [v for v in [buscar_valor_da_coluna(l, nome_col) for l in swot] if v and str(v).strip()]
        contagens[rep_name] = len(_aplicar_limite(vals, limite))

    inv = ler_guia_de_arquivo_xlsx(caminho_xlsx, "3.2. Inventário de necessidades", list(COLUNAS_INVENTARIO_NECESSIDADES))
    contagens["needs"] = len(_aplicar_limite(inv, limite))

    metas = ler_guia_de_arquivo_xlsx(caminho_xlsx, "4.1. Metas e Ações", list(COLUNAS_METAS_ACOES))
    contagens["goals-actions"] = len(_aplicar_limite(metas, limite))
    contagens["kpis"] = len(_aplicar_limite(metas, limite))

    orc = ler_guia_de_arquivo_xlsx(caminho_xlsx, "4.2. Plano Orçamentário", list(COLUNAS_PLANO_ORCAMENTARIO))
    contagens["budget-plan"] = len(_aplicar_limite(orc, limite))

    riscos = ler_guia_de_arquivo_xlsx(caminho_xlsx, "4.3. Gestão de Riscos", list(COLUNAS_GESTAO_RISCOS))
    contagens["risk-plan"] = len(_aplicar_limite(riscos, limite))

    cap = ler_guia_de_arquivo_xlsx(caminho_xlsx, "5.1. Plano de Capacitação", list(COLUNAS_CAPACITACAO))
    for nome_col, rep_name in [
        ("TREINAMENTOS TÉCNICOS", "technical-trainings"),
        ("AÇÕES DE RETENÇÃO DE TALENTOS", "talent-actions"),
    ]:
        vals = [v for v in [buscar_valor_da_coluna(l, nome_col) for l in cap] if v and str(v).strip()]
        contagens[rep_name] = len(_aplicar_limite(vals, limite))

    acoes = ler_guia_de_arquivo_xlsx(caminho_xlsx, "5.3. Ações Específicas", list(COLUNAS_ACOES_ESPECIFICAS))
    for nome_col, rep_name in [
        ("AÇÕES MACRO", "macro-actions"),
        ("AÇÕES MICRO", "micro-actions"),
        ("AÇÕES EXTRAS", "extra-actions"),
    ]:
        vals = [v for v in [buscar_valor_da_coluna(l, nome_col) for l in acoes] if v and str(v).strip()]
        contagens[rep_name] = len(_aplicar_limite(vals, limite))

    return contagens


def executar_analise_completa_campos(
    ctrl: ControladorExecucaoOculta,
    caminho_xlsx: Optional[Path] = None,
    limite: Optional[int] = None,
) -> Dict[str, Any]:
    """
    Executa a análise completa da tela atual do Nex Privacy vs Planilha:
    - Quantos cards existem, quantos estão preenchidos, vazios ou parciais (ex: parou no meio de um card!)
    - Quais são os nomes e tipos reais dos campos HTML de cada bloco (ex: Plano Orçamentário)
    - Quantas variáveis `[name]` existem no `<form>` atual e quantas existirão após o preenchimento (alerta `max_input_vars = 1000`)
    """
    log_debug("ANÁLISE DE CAMPOS", "Inspecionando todos os blocos, seletores HTML e capacidade do DOM na guia aberta...", nivel="INFO")
    diag = ctrl.inspecionar_dom(timeout=20.0)
    if not diag.get("ok"):
        log_debug("ANÁLISE DE CAMPOS", f"Não foi possível inspecionar o DOM: {diag.get('erro')}", nivel="WARN")
        return diag

    contagens_planilha = _contar_linhas_planilha_por_repeater(caminho_xlsx, limite) if caminho_xlsx else {}

    DADOS_DIR.mkdir(parents=True, exist_ok=True)
    arq_json = DADOS_DIR / "diagnostico_campos_dom.json"
    arq_json.write_text(json.dumps(diag, ensure_ascii=False, indent=2), encoding="utf-8")

    form_info = diag.get("formInfo") or {}
    reps = diag.get("repeaters") or {}

    print("\n" + "=" * 100)
    print("🔬 DIAGNÓSTICO COMPLETO DOS CAMPOS DO SISTEMA (DOM ATUAL × PLANILHA)")
    print("=" * 100)
    print(f"🌐 URL da Guia:       {diag.get('url')}")
    print(f"📍 Etapa Ativa:       {diag.get('etapaAtiva')}")
    if form_info.get("existe"):
        print(
            f"📝 Formulário HTML:   method={form_info.get('method')} | enctype={form_info.get('enctype')} | "
            f"Total atual de variáveis [name] no <form> = {form_info.get('totalNamedFields')}"
        )
        janela1000 = form_info.get("janelaLimite1000") or []
        if janela1000:
            print("🚨 JANELA EM TORNO DA 1.000ª VARIÁVEL DO <form> (Limite padrão PHP max_input_vars = 1000):")
            for item_j in janela1000:
                print(
                    f"   -> Posição #{item_j.get('pos1Based')}: <{item_j.get('tag')} name=\"{item_j.get('name')}\"> "
                    f"= \"{item_j.get('valorPreview')}\""
                )
    print("-" * 100)
    print(f"{'BLOCO / REPEATER':<38} | {'PLANILHA':>8} | {'DOM ATUAL':>9} | {'PREENCH.':>8} | {'PARCIAIS':>8} | {'+CRIAR':>7} | {'CAMPOS/CARD':>11} | {'BOTÃO +ADD'}")
    print("-" * 100)

    inputs_atuais_repeaters = 0
    inputs_projetados_repeaters = 0

    for rep_key, info in reps.items():
        lbl = f"{info.get('label')} ({rep_key})"
        qtd_plan = contagens_planilha.get(rep_key, -1)
        str_plan = str(qtd_plan) if qtd_plan >= 0 else "-"
        if not info.get("exists"):
            print(f"{lbl[:38]:<38} | {str_plan:>8} | {'AUSENTE':>9} | {'-':>8} | {'-':>8} | {'-':>7} | {'-':>11} | ❌")
            continue

        tot_r = int(info.get("totalRows", 0))
        fill_r = int(info.get("filledRows", 0))
        part_r = int(info.get("qtdPartialRows", 0))
        schema = info.get("schema") or []
        qtd_campos_card = len(schema)
        a_criar = max(0, qtd_plan - tot_r) if qtd_plan >= 0 else 0
        btn_txt = "OK ✅" if (info.get("addBtnExists") and not info.get("addBtnDisabled")) else "❌"

        inputs_atuais_repeaters += tot_r * max(1, qtd_campos_card)
        if qtd_plan >= 0:
            inputs_projetados_repeaters += max(tot_r, qtd_plan) * max(1, qtd_campos_card)

        print(
            f"{lbl[:38]:<38} | {str_plan:>8} | {tot_r:>9} | {fill_r:>8} | {part_r:>8} | {('+' + str(a_criar)):>7} | {qtd_campos_card:>11} | {btn_txt}"
        )

    print("-" * 100)
    # Detalhar quaisquer blocos com cards parciais ou diferenças relevantes e exibir o schema de 4.1 e 4.2
    for rep_key, info in reps.items():
        if not info.get("exists"):
            continue
        parciais = info.get("partialRows") or []
        if parciais:
            print(f"⚠️  [{info.get('label')}] {info.get('qtdPartialRows')} card(s) PARCIALMENTE preenchido(s) na tela atual:")
            for p in parciais[:5]:
                print(
                    f"    • Card #{p.get('card')}: {p.get('qtdPreenchidos')}/{p.get('totalCampos')} preenchidos -> "
                    f"Preenchidos={p.get('preenchidos')} | Vazios={p.get('vazios')}"
                )
        dups = int(info.get("duplicateNamesCount", 0))
        if dups > 0:
            print(f"⚠️  [{info.get('label')}] {dups} atributo(s) `name` duplicado(s) no DOM: {info.get('duplicateNamesSample')}")

        if rep_key in ("goals-actions", "budget-plan", "risk-plan", "needs"):
            schema = info.get("schema") or []
            resumo_s = " | ".join(
                f"[{c.get('idx')}] <{c.get('tag')} type='{c.get('type')}' name='{c.get('name')}' label='{c.get('label')}'{' readonly' if c.get('readonly') else ''}>"
                for c in schema
            )
            print(f"🔎 Schema [{info.get('label')}]: {resumo_s}")

    if form_info.get("existe") and contagens_planilha:
        fixos_estimados = max(0, int(form_info.get("totalNamedFields", 0)) - inputs_atuais_repeaters)
        total_final_estimado = fixos_estimados + inputs_projetados_repeaters
        print("-" * 100)
        print(
            f"📐 PROJEÇÃO DE VARIÁVEIS POST NO <form>: ~{total_final_estimado} inputs com atributo `name` "
            f"(fixos={fixos_estimados} + repeaters={inputs_projetados_repeaters})."
        )
        if total_final_estimado > 1000:
            print(
                "🚨 ATENÇÃO CRÍTICA: O formulário completo gera mais de 1.000 campos HTML (`name=\"...\"`).\n"
                "   Se o servidor PHP do Nex Privacy estiver com `max_input_vars = 1000` (padrão do PHP),\n"
                "   ao clicar em 'Salvar' o servidor descartará silenciosamente todos os campos após o #1.000!"
            )
    print(f"💾 JSON completo da inspeção salvo em: {arq_json}")
    print("=" * 100 + "\n", flush=True)
    return diag


def executar_automacao_segundo_plano(
    alvo: str = "todas",
    caminho_planilha: Optional[str] = None,
    limite: Optional[int] = None,
    porta_cdp: int = PORTA_PONTE_LOCAL,
    timeout_conexao: int = TIMEOUT_CONEXAO_PADRAO,
    delay_add_ms: int = DELAY_ADICIONAR_MS_PADRAO,
    tamanho_lote: int = TAMANHO_LOTE_PADRAO,
    limpar_antes: bool = False,
    apenas_analisar: bool = False,
) -> bool:
    """
    Executa a automação 100% oculta com `[DEBUG]` passo a passo, análise prévia dos campos do DOM,
    lotes cadenciados com maior tempo de espera e verificação por leitura reversa.
    """
    chave = str(alvo or "todas").strip().lower()
    mapeamento = carregar_mapeamento() or {}

    log_debug(
        "INÍCIO",
        f"Alvo='{chave}' | Limite={limite or 'SEM LIMITE'} | DelayAdd={delay_add_ms}ms | "
        f"Lote={tamanho_lote} | LimparAntes={limpar_antes} | ApenasAnalisar={apenas_analisar} | "
        f"Timeout Conexão={timeout_conexao}s",
        nivel="INFO",
    )

    ctrl = ControladorExecucaoOculta(
        porta_ponte=porta_cdp,
        porta_cdp=PORTA_CDP_PADRAO,
        delay_add_ms=delay_add_ms,
        tamanho_lote=tamanho_lote,
        limpar_antes=limpar_antes,
    )
    try:
        # PASSO 1: Conectar à guia já aberta do Nex Privacy (com Timer regressivo e diagnóstico)
        if not conectar_guia_com_debug_e_timer(ctrl, timeout_segundos=timeout_conexao):
            return False

        # PASSO 2: Obter e validar a planilha
        caminho_xlsx = obter_planilha_com_debug_e_timer(ctrl, caminho_planilha)
        if caminho_xlsx is None and not apenas_analisar and chave not in ("pagina6", "p6", "6", "6.1", "kpis"):
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

        # PASSO 2.5: Análise completa dos campos no DOM vs linhas da planilha
        executar_analise_completa_campos(ctrl, caminho_xlsx, limite)
        if apenas_analisar or chave in ("analisar", "diagnostico", "inspecionar"):
            log_debug("ANÁLISE CONCLUÍDA", "Modo somente-análise finalizado sem alterar os campos da tela.", nivel="OK")
            return True

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
                "🎉 Todas as etapas solicitadas foram preenchidas em 2º plano com sucesso!",
                nivel="OK",
            )
        return ok
    finally:
        # Gerar o relatório de auditoria antes de encerrar
        if hasattr(ctrl, 'historico_relatorio') and ctrl.historico_relatorio:
            caminho_relatorio = ctrl.gerar_relatorio_auditoria()
            if caminho_relatorio:
                print("\n" + "=" * 78)
                print(f"📑 RELATÓRIO DE AUDITORIA GERADO!")
                print(f"📁 Salvo em: {caminho_relatorio}")
                print(f"Você pode abrir este arquivo para conferir o que foi digitado em cada campo.")
                print("=" * 78 + "\n")
                
        ctrl.fechar()
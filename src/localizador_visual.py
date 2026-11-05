"""
Localizador visual inteligente da página (Win32 GDI + OpenCV + Ctrl+F):
1. Captura a tela via API nativa Win32 GDI (imune a falhas do ImageGrab em múltiplos monitores).
2. Usa Ctrl+F para destacar o nome da coluna/campo na página.
3. Com OpenCV, valida qual destaque está pareado na mesma linha horizontal com um botão
   azul sólido 'Adicionar' (#2C49F5) ou se é um campo simples abaixo do rótulo,
   ignorando ocorrências dentro de frases ou da caixa 'Dicas e exemplos'.
"""

import base64
import ctypes
from ctypes import wintypes
from html.parser import HTMLParser
from pathlib import Path
import re
import subprocess
import time
from typing import Any, Dict, List, Optional, Tuple
import unicodedata

import cv2
import numpy as np
from PIL import ImageGrab

from src.config import BASE_DIR, DADOS_DIR, ESPERA_APOS_ADICIONAR, carregar_mapeamento
from src.driver_windows import (
    atalho_teclado,
    copiar_para_clipboard,
    gdi32,
    mover_e_clicar,
    pressionar_tecla_unica,
    rolar_pagina,
    user32,
)

MAPA_TITULO_PAGINA = {
    "FORÇA": "Forças",
    "FORCAS": "Forças",
    "FORÇAS": "Forças",
    "FRAQUEZA": "Fraquezas",
    "FRAQUEZAS": "Fraquezas",
    "OPORTUNIDADE": "Oportunidades",
    "OPORTUNIDADES": "Oportunidades",
    "AMEAÇA": "Ameaças",
    "AMEACA": "Ameaças",
    "AMEACAS": "Ameaças",
    "AMEAÇAS": "Ameaças",
}

# Regras estruturais extraídas do DOM HTML (.wizard-table-card / data-repeater / Bootstrap grid)
# para limitar a atuação de cada bloco à sua própria <div> e diferenciar os botões Adicionar:
REGRAS_DIV_REPEATER_HTML: Dict[str, Dict[str, Any]] = {
    "FORCAS": {
        "data_repeater": "strengths",
        "data_repeater_add": "strengths",
        "texto_botao": "Adicionar",
        "coluna_grid": "col-md-6-esq",
        "x_min_input": 655,
        "x_max_input": 1165,
        "x_min_btn_add": 1100,
        "x_max_btn_add": 1235,
        "w_min_btn_add": 52,
        "w_max_btn_add": 108,
        "tem_bloco_abaixo_mesma_coluna": True,
    },
    "FRAQUEZAS": {
        "data_repeater": "weaknesses",
        "data_repeater_add": "weaknesses",
        "texto_botao": "Adicionar",
        "coluna_grid": "col-md-6-dir",
        "x_min_input": 1255,
        "x_max_input": 1765,
        "x_min_btn_add": 1700,
        "x_max_btn_add": 1840,
        "w_min_btn_add": 52,
        "w_max_btn_add": 108,
        "tem_bloco_abaixo_mesma_coluna": True,
    },
    "OPORTUNIDADES": {
        "data_repeater": "opportunities",
        "data_repeater_add": "opportunities",
        "texto_botao": "Adicionar",
        "coluna_grid": "col-md-6-esq",
        "x_min_input": 655,
        "x_max_input": 1165,
        "x_min_btn_add": 1100,
        "x_max_btn_add": 1235,
        "w_min_btn_add": 52,
        "w_max_btn_add": 108,
        "tem_bloco_abaixo_mesma_coluna": False,
        "bloqueado_por_botao_largo_abaixo": True,
    },
    "AMEACAS": {
        "data_repeater": "threats",
        "data_repeater_add": "threats",
        "texto_botao": "Adicionar",
        "coluna_grid": "col-md-6-dir",
        "x_min_input": 1255,
        "x_max_input": 1765,
        "x_min_btn_add": 1700,
        "x_max_btn_add": 1840,
        "w_min_btn_add": 52,
        "w_max_btn_add": 108,
        "tem_bloco_abaixo_mesma_coluna": False,
        "bloqueado_por_botao_largo_abaixo": True,
    },
    "DESCRICAO": {
        "data_repeater": "needs",
        "data_repeater_add": "needs",
        "texto_botao": "Adicionar necessidade",
        "coluna_grid": "col-md-6",
        "x_min_input": 655,
        "x_max_input": 1240,
        "w_min_btn_add": 115,
        "w_max_btn_add": 230,
        "ultimo_bloco_da_pagina": True,
    },
    "PRIORIDADE": {
        "data_repeater": "needs",
        "data_repeater_add": "needs",
        "texto_botao": "Adicionar necessidade",
        "coluna_grid": "col-md-2",
        "x_min_input": 1245,
        "x_max_input": 1455,
        "w_min_btn_add": 115,
        "w_max_btn_add": 230,
        "ultimo_bloco_da_pagina": True,
    },
    "AREA DEMANDANTE": {
        "data_repeater": "needs",
        "data_repeater_add": "needs",
        "texto_botao": "Adicionar necessidade",
        "coluna_grid": "col-md-3",
        "x_min_input": 1455,
        "x_max_input": 1745,
        "w_min_btn_add": 115,
        "w_max_btn_add": 230,
        "ultimo_bloco_da_pagina": True,
    },
    "JUSTIFICATIVA": {
        "data_repeater": "needs",
        "data_repeater_add": "needs",
        "texto_botao": "Adicionar necessidade",
        "coluna_grid": "col-12",
        "x_min_input": 655,
        "x_max_input": 1835,
        "w_min_btn_add": 115,
        "w_max_btn_add": 230,
        "ultimo_bloco_da_pagina": True,
    },
    # =========================================================================
    # PÁGINA 4 ('4. Metas e Planos Obrigatórios') — 3 BLOCOS (.wizard-table-card)
    # Bloco 1: 'Metas e ações' (data-repeater="goals-actions", botão "Adicionar ação")
    # =========================================================================
    "ID": {
        "data_repeater": "goals-actions",
        "data_repeater_add": "goals-actions",
        "texto_botao": "Adicionar ação",
        "coluna_grid": "col-md-2",
        "x_min_input": 675,
        "x_max_input": 855,
        "w_min_btn_add": 95,
        "w_max_btn_add": 185,
    },
    "OBJETIVO ESTRATEGICO": {
        "data_repeater": "goals-actions",
        "data_repeater_add": "goals-actions",
        "texto_botao": "Adicionar ação",
        "coluna_grid": "col-md-5",
        "x_min_input": 870,
        "x_max_input": 1330,
        "w_min_btn_add": 95,
        "w_max_btn_add": 185,
    },
    "META": {
        "data_repeater": "goals-actions",
        "data_repeater_add": "goals-actions",
        "texto_botao": "Adicionar ação",
        "coluna_grid": "col-md-5",
        "x_min_input": 1345,
        "x_max_input": 1805,
        "w_min_btn_add": 95,
        "w_max_btn_add": 185,
    },
    "ACAO PRINCIPAL": {
        "data_repeater": "goals-actions",
        "data_repeater_add": "goals-actions",
        "texto_botao": "Adicionar ação",
        "coluna_grid": "col-md-4",
        "x_min_input": 675,
        "x_max_input": 1045,
        "w_min_btn_add": 95,
        "w_max_btn_add": 185,
    },
    "INDICADOR": {
        "data_repeater": "goals-actions",
        "data_repeater_add": "goals-actions",
        "texto_botao": "Adicionar ação",
        "coluna_grid": "col-md-3",
        "x_min_input": 675,
        "x_max_input": 1330,
        "w_min_btn_add": 95,
        "w_max_btn_add": 185,
    },
    "PRAZO": {
        "data_repeater": "goals-actions",
        "data_repeater_add": "goals-actions",
        "texto_botao": "Adicionar ação",
        "coluna_grid": "col-md-3",
        "x_min_input": 1345,
        "x_max_input": 1615,
        "w_min_btn_add": 95,
        "w_max_btn_add": 185,
    },
    "RESPONSAVEL": {
        "data_repeater": "goals-actions",
        "data_repeater_add": "goals-actions",
        "texto_botao": "Adicionar ação",
        "coluna_grid": "col-md-12",
        "x_min_input": 675,
        "x_max_input": 1805,
        "w_min_btn_add": 95,
        "w_max_btn_add": 185,
    },
    # =========================================================================
    # Bloco 2: 'Plano orçamentário e de investimentos' (data-repeater="budget-plan", botão "Adicionar linha")
    # =========================================================================
    "ACAO (ID)": {
        "data_repeater": "budget-plan",
        "data_repeater_add": "budget-plan",
        "texto_botao": "Adicionar linha",
        "coluna_grid": "col-md-2",
        "x_min_input": 675,
        "x_max_input": 840,
        "w_min_btn_add": 95,
        "w_max_btn_add": 185,
    },
    "DESCRICAO DO CUSTO": {
        "data_repeater": "budget-plan",
        "data_repeater_add": "budget-plan",
        "texto_botao": "Adicionar linha",
        "coluna_grid": "col-md-6",
        "x_min_input": 855,
        "x_max_input": 1360,
        "w_min_btn_add": 95,
        "w_max_btn_add": 185,
    },
    "FONTE DE RECURSO": {
        "data_repeater": "budget-plan",
        "data_repeater_add": "budget-plan",
        "texto_botao": "Adicionar linha",
        "coluna_grid": "col-md-4",
        "x_min_input": 1375,
        "x_max_input": 1710,
        "w_min_btn_add": 95,
        "w_max_btn_add": 185,
    },
    "ANO 1": {
        "data_repeater": "budget-plan",
        "data_repeater_add": "budget-plan",
        "texto_botao": "Adicionar linha",
        "coluna_grid": "col-md-3",
        "x_min_input": 675,
        "x_max_input": 925,
        "w_min_btn_add": 95,
        "w_max_btn_add": 185,
    },
    "ANO 2": {
        "data_repeater": "budget-plan",
        "data_repeater_add": "budget-plan",
        "texto_botao": "Adicionar linha",
        "coluna_grid": "col-md-3",
        "x_min_input": 940,
        "x_max_input": 1185,
        "w_min_btn_add": 95,
        "w_max_btn_add": 185,
    },
    "ANO 3": {
        "data_repeater": "budget-plan",
        "data_repeater_add": "budget-plan",
        "texto_botao": "Adicionar linha",
        "coluna_grid": "col-md-3",
        "x_min_input": 1200,
        "x_max_input": 1445,
        "w_min_btn_add": 95,
        "w_max_btn_add": 185,
    },
    "ANO 4": {
        "data_repeater": "budget-plan",
        "data_repeater_add": "budget-plan",
        "texto_botao": "Adicionar linha",
        "coluna_grid": "col-md-3",
        "x_min_input": 1460,
        "x_max_input": 1710,
        "w_min_btn_add": 95,
        "w_max_btn_add": 185,
    },
    # =========================================================================
    # Bloco 3: 'Plano de gestão de riscos' (data-repeater="risk-plan", botão "Adicionar risco")
    # =========================================================================
    "RISCO IDENTIFICADO": {
        "data_repeater": "risk-plan",
        "data_repeater_add": "risk-plan",
        "texto_botao": "Adicionar risco",
        "coluna_grid": "col-md-4",
        "x_min_input": 675,
        "x_max_input": 1045,
        "w_min_btn_add": 95,
        "w_max_btn_add": 185,
        "ultimo_bloco_da_pagina": True,
    },
    "PROBABILIDADE": {
        "data_repeater": "risk-plan",
        "data_repeater_add": "risk-plan",
        "texto_botao": "Adicionar risco",
        "coluna_grid": "col-md-2",
        "x_min_input": 1060,
        "x_max_input": 1235,
        "w_min_btn_add": 95,
        "w_max_btn_add": 185,
        "ultimo_bloco_da_pagina": True,
    },
    "IMPACTO": {
        "data_repeater": "risk-plan",
        "data_repeater_add": "risk-plan",
        "texto_botao": "Adicionar risco",
        "coluna_grid": "col-md-2",
        "x_min_input": 1250,
        "x_max_input": 1425,
        "w_min_btn_add": 95,
        "w_max_btn_add": 185,
        "ultimo_bloco_da_pagina": True,
    },
    "NIVEL": {
        "data_repeater": "risk-plan",
        "data_repeater_add": "risk-plan",
        "texto_botao": "Adicionar risco",
        "coluna_grid": "col-md-2",
        "x_min_input": 1440,
        "x_max_input": 1615,
        "w_min_btn_add": 95,
        "w_max_btn_add": 185,
        "ultimo_bloco_da_pagina": True,
    },
    "ACAO DE MITIGACAO": {
        "data_repeater": "risk-plan",
        "data_repeater_add": "risk-plan",
        "texto_botao": "Adicionar risco",
        "coluna_grid": "col-md-8",
        "x_min_input": 675,
        "x_max_input": 1425,
        "w_min_btn_add": 95,
        "w_max_btn_add": 185,
        "ultimo_bloco_da_pagina": True,
    },
    "RESPONSAVEL PELO RISCO": {
        "data_repeater": "risk-plan",
        "data_repeater_add": "risk-plan",
        "texto_botao": "Adicionar risco",
        "coluna_grid": "col-md-4",
        "x_min_input": 1440,
        "x_max_input": 1805,
        "w_min_btn_add": 95,
        "w_max_btn_add": 185,
        "ultimo_bloco_da_pagina": True,
    },
    # =========================================================================
    # Página 5: '5. Capacitação, Comunicação e Ações Específicas'
    # Passo A (2 colunas col-lg-6) + Passo B (3 colunas col-lg-4)
    # =========================================================================
    "TREINAMENTOS TECNICOS": {
        "data_repeater": "technical-trainings",
        "data_repeater_add": "technical-trainings",
        "texto_botao": "Adicionar",
        "coluna_grid": "col-lg-6",
        "x_min_input": 675,
        "x_max_input": 1185,
        "w_min_btn_add": 68,
        "w_max_btn_add": 112,
    },
    "ACOES DE RETENCAO DE TALENTOS": {
        "data_repeater": "talent-actions",
        "data_repeater_add": "talent-actions",
        "texto_botao": "Adicionar",
        "coluna_grid": "col-lg-6",
        "x_min_input": 1245,
        "x_max_input": 1795,
        "w_min_btn_add": 68,
        "w_max_btn_add": 112,
    },
    "ACOES MACRO": {
        "data_repeater": "macro-actions",
        "data_repeater_add": "macro-actions",
        "texto_botao": "Adicionar",
        "coluna_grid": "col-lg-4",
        "x_min_input": 675,
        "x_max_input": 1005,
        "w_min_btn_add": 68,
        "w_max_btn_add": 112,
        "ultimo_bloco_da_pagina": True,
    },
    "ACOES MICRO": {
        "data_repeater": "micro-actions",
        "data_repeater_add": "micro-actions",
        "texto_botao": "Adicionar",
        "coluna_grid": "col-lg-4",
        "x_min_input": 1055,
        "x_max_input": 1415,
        "w_min_btn_add": 68,
        "w_max_btn_add": 112,
        "ultimo_bloco_da_pagina": True,
    },
    "ACOES EXTRAS": {
        "data_repeater": "extra-actions",
        "data_repeater_add": "extra-actions",
        "texto_botao": "Adicionar",
        "coluna_grid": "col-lg-4",
        "x_min_input": 1465,
        "x_max_input": 1805,
        "w_min_btn_add": 68,
        "w_max_btn_add": 112,
        "ultimo_bloco_da_pagina": True,
    },
    # =========================================================================
    # Página 6: '6. Monitoramento, Anexos e Revisão' — 'Indicadores de sucesso (KPIs)'
    # (data-repeater="kpis", botão "Adicionar KPI", inicia com 3 linhas vazias no HTML)
    # =========================================================================
    "FORMULA / DESCRICAO": {
        "data_repeater": "kpis",
        "data_repeater_add": "kpis",
        "texto_botao": "Adicionar KPI",
        "coluna_grid": "col-md-6",
        "x_min_input": 1125,
        "x_max_input": 1710,
        "w_min_btn_add": 90,
        "w_max_btn_add": 175,
        "ultimo_bloco_da_pagina": True,
    },
    "FORMULA/DESCRICAO": {
        "data_repeater": "kpis",
        "data_repeater_add": "kpis",
        "texto_botao": "Adicionar KPI",
        "coluna_grid": "col-md-6",
        "x_min_input": 1125,
        "x_max_input": 1710,
        "w_min_btn_add": 90,
        "w_max_btn_add": 175,
        "ultimo_bloco_da_pagina": True,
    },
}

SRCCOPY = 0x00CC0020
CAPTUREBLT = 0x40000000
DIB_RGB_COLORS = 0
ARQUIVO_LOG_VISAO = DADOS_DIR / "debug_visao.log"


def _log_visao(msg: str) -> None:
    try:
        DADOS_DIR.mkdir(parents=True, exist_ok=True)
        ts = time.strftime("%H:%M:%S")
        with open(ARQUIVO_LOG_VISAO, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {msg}\n")
    except Exception:
        pass


def _norm_chave(texto: str) -> str:
    sem_acento = "".join(
        c for c in unicodedata.normalize("NFD", str(texto or "")) if unicodedata.category(c) != "Mn"
    )
    return re.sub(r"\s+", " ", sem_acento).strip().upper()


_normalizar_chave = _norm_chave


class _ExtratorCamposHTML(HTMLParser):
    """Extrai automaticamente do HTML os títulos de cards (h5) com botão Adicionar e rótulos (label)."""

    def __init__(self) -> None:
        super().__init__()
        self.tag_atual: Optional[str] = None
        self.buffer_texto: List[str] = []
        self.ultimo_h5: Optional[str] = None
        self.campos: Dict[str, Dict[str, Any]] = {}
        self.ordem_h5 = 0

    def handle_starttag(self, tag: str, attrs: List[Tuple[str, Optional[str]]]) -> None:
        attrs_dict = dict(attrs)
        if tag in ("h5", "label"):
            self.tag_atual = tag
            self.buffer_texto = []
        elif tag == "button" and "data-repeater-add" in attrs_dict and self.ultimo_h5:
            chave = _norm_chave(self.ultimo_h5)
            if chave in self.campos:
                self.campos[chave]["tem_botao_adicionar"] = True
                self.campos[chave]["repeater_id"] = attrs_dict.get("data-repeater-add") or ""

    def handle_data(self, data: str) -> None:
        if self.tag_atual in ("h5", "label"):
            self.buffer_texto.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag == self.tag_atual:
            texto = re.sub(r"\s+", " ", "".join(self.buffer_texto)).strip()
            self.tag_atual = None
            self.buffer_texto = []
            if not texto or len(texto) > 65:
                return
            if texto.lower() in ("orientação geral", "dicas e exemplos", "notificações"):
                return
            chave = _norm_chave(texto)
            if tag == "h5":
                self.ultimo_h5 = texto
                if chave not in self.campos:
                    self.campos[chave] = {
                        "titulo_exato": texto,
                        "tem_botao_adicionar": False,
                        "tipo": "card",
                        "ordem": self.ordem_h5,
                    }
                    self.ordem_h5 += 1
            elif tag == "label" and chave not in self.campos:
                self.campos[chave] = {
                    "titulo_exato": texto,
                    "tem_botao_adicionar": False,
                    "tipo": "label",
                    "ordem": 999,
                }


_CACHE_MAPA_HTML: Optional[Dict[str, Dict[str, Any]]] = None


def obter_mapa_html_campos() -> Dict[str, Dict[str, Any]]:
    """Lê qualquer arquivo .html salvo na pasta dados/ para enriquecer a busca com os nomes exatos do DOM."""
    global _CACHE_MAPA_HTML
    if _CACHE_MAPA_HTML is not None:
        return _CACHE_MAPA_HTML

    mapa: Dict[str, Dict[str, Any]] = {}
    try:
        arquivos_html = sorted(DADOS_DIR.glob("*.html")) + sorted(BASE_DIR.glob("*.html"))
        for arq in arquivos_html:
            conteudo = arq.read_text(encoding="utf-8", errors="ignore")
            parser = _ExtratorCamposHTML()
            parser.feed(conteudo)
            for k, v in parser.campos.items():
                if k not in mapa or v.get("tem_botao_adicionar"):
                    mapa[k] = v
                if k.endswith("S") and len(k) > 4:
                    k_sing = k[:-1]
                    if k_sing not in mapa:
                        mapa[k_sing] = v
    except Exception as exc:
        _log_visao(f"Aviso ao ler HTML local: {exc}")

    _CACHE_MAPA_HTML = mapa
    return mapa


def obter_parametros_calibrados(mapeamento: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """
    Retorna os parâmetros geométricos relativos (distâncias entre título, caixa de texto e botão Adicionar)
    a partir de 1 único conjunto de coordenadas calibrado pelo usuário.
    """
    if mapeamento is None:
        mapeamento = carregar_mapeamento()

    params = dict(mapeamento.get("parametros_referencia") or {})
    campos = mapeamento.get("campos_detalhados") or {}

    ponto_planilha = params.get("planilha")
    if not ponto_planilha:
        for info in campos.values():
            if "planilha_linha1" in info:
                ponto_planilha = info["planilha_linha1"]
                break
    if not ponto_planilha:
        ponto_planilha = [-1649, 373]

    # Distâncias relativas padrão baseadas no layout CSS (.wizard-table-card / Bootstrap .col-md-6)
    dx_botao_para_campo = int(params.get("dx_botao_para_campo", -250))
    dy_botao_para_campo = int(params.get("dy_botao_para_campo", 72))
    dx_titulo_para_campo = int(params.get("dx_titulo_para_campo", 240))
    dy_titulo_para_campo = int(params.get("dy_titulo_para_campo", 72))
    dx_titulo_para_botao = int(params.get("dx_titulo_para_botao", 555))
    dy_titulo_para_botao = int(params.get("dy_titulo_para_botao", 0))
    passo_y_nova_caixa = int(params.get("passo_y_nova_caixa", 84))

    # Garante limites seguros para nunca clicar fora da caixa <input>
    dx_botao_para_campo = max(-350, min(-140, dx_botao_para_campo))
    dy_botao_para_campo = max(56, min(92, dy_botao_para_campo))
    dx_titulo_para_campo = max(80, min(380, dx_titulo_para_campo))
    dy_titulo_para_campo = max(56, min(92, dy_titulo_para_campo))
    passo_y_nova_caixa = max(68, min(98, passo_y_nova_caixa))

    ponto_neutro = params.get("ponto_neutro", [620, 200])

    return {
        "planilha": [int(ponto_planilha[0]), int(ponto_planilha[1])],
        "ponto_neutro": [int(ponto_neutro[0]), int(ponto_neutro[1])],
        "dx_botao_para_campo": dx_botao_para_campo,
        "dy_botao_para_campo": dy_botao_para_campo,
        "dx_titulo_para_campo": dx_titulo_para_campo,
        "dy_titulo_para_campo": dy_titulo_para_campo,
        "dx_titulo_para_botao": dx_titulo_para_botao,
        "dy_titulo_para_botao": dy_titulo_para_botao,
        "passo_y_nova_caixa": passo_y_nova_caixa,
        "titulo_ref": params.get("titulo_ref"),
        "campo_ref": params.get("campo_ref"),
        "botao_ref": params.get("botao_ref"),
    }


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ("biSize", wintypes.DWORD),
        ("biWidth", wintypes.LONG),
        ("biHeight", wintypes.LONG),
        ("biPlanes", wintypes.WORD),
        ("biBitCount", wintypes.WORD),
        ("biCompression", wintypes.DWORD),
        ("biSizeImage", wintypes.DWORD),
        ("biXPelsPerMeter", wintypes.LONG),
        ("biYPelsPerMeter", wintypes.LONG),
        ("biClrUsed", wintypes.DWORD),
        ("biClrImportant", wintypes.DWORD),
    ]


user32.GetDC.argtypes = [wintypes.HWND]
user32.GetDC.restype = wintypes.HDC
user32.ReleaseDC.argtypes = [wintypes.HWND, wintypes.HDC]
user32.ReleaseDC.restype = ctypes.c_int

gdi32.CreateCompatibleDC.argtypes = [wintypes.HDC]
gdi32.CreateCompatibleDC.restype = wintypes.HDC
gdi32.CreateCompatibleBitmap.argtypes = [wintypes.HDC, ctypes.c_int, ctypes.c_int]
gdi32.CreateCompatibleBitmap.restype = wintypes.HBITMAP
gdi32.SelectObject.argtypes = [wintypes.HDC, wintypes.HGDIOBJ]
gdi32.SelectObject.restype = wintypes.HGDIOBJ
gdi32.BitBlt.argtypes = [
    wintypes.HDC,
    ctypes.c_int,
    ctypes.c_int,
    ctypes.c_int,
    ctypes.c_int,
    wintypes.HDC,
    ctypes.c_int,
    ctypes.c_int,
    wintypes.DWORD,
]
gdi32.BitBlt.restype = wintypes.BOOL
gdi32.GetDIBits.argtypes = [
    wintypes.HDC,
    wintypes.HBITMAP,
    wintypes.UINT,
    wintypes.UINT,
    wintypes.LPVOID,
    ctypes.POINTER(BITMAPINFOHEADER),
    wintypes.UINT,
]
gdi32.GetDIBits.restype = ctypes.c_int
gdi32.DeleteObject.argtypes = [wintypes.HGDIOBJ]
gdi32.DeleteObject.restype = wintypes.BOOL
gdi32.DeleteDC.argtypes = [wintypes.HDC]
gdi32.DeleteDC.restype = wintypes.BOOL


def _capturar_regiao_gdi(left: int, top: int, width: int, height: int) -> Optional[np.ndarray]:
    """Captura uma região da tela diretamente via Win32 GDI (suporta múltiplos monitores)."""
    if width <= 0 or height <= 0:
        return None
    hdc_screen = user32.GetDC(None)
    if not hdc_screen:
        return None
    hdc_mem = gdi32.CreateCompatibleDC(hdc_screen)
    if not hdc_mem:
        user32.ReleaseDC(None, hdc_screen)
        return None
    hbmp = gdi32.CreateCompatibleBitmap(hdc_screen, width, height)
    if not hbmp:
        gdi32.DeleteDC(hdc_mem)
        user32.ReleaseDC(None, hdc_screen)
        return None

    old_obj = gdi32.SelectObject(hdc_mem, hbmp)
    try:
        ok = gdi32.BitBlt(
            hdc_mem,
            0,
            0,
            width,
            height,
            hdc_screen,
            left,
            top,
            SRCCOPY | CAPTUREBLT,
        )
        if not ok:
            return None

        bmi = BITMAPINFOHEADER()
        bmi.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        bmi.biWidth = width
        bmi.biHeight = -height
        bmi.biPlanes = 1
        bmi.biBitCount = 32
        bmi.biCompression = 0

        buf = ctypes.create_string_buffer(width * height * 4)
        linhas = gdi32.GetDIBits(
            hdc_mem,
            hbmp,
            0,
            height,
            buf,
            ctypes.byref(bmi),
            DIB_RGB_COLORS,
        )
        if linhas <= 0:
            return None

        arr = np.frombuffer(buf, dtype=np.uint8).reshape((height, width, 4))
        return arr[:, :, :3].copy()
    finally:
        gdi32.SelectObject(hdc_mem, old_obj)
        gdi32.DeleteObject(hbmp)
        gdi32.DeleteDC(hdc_mem)
        user32.ReleaseDC(None, hdc_screen)


def capturar_tela_sistema(x_ref: int = 1000) -> Tuple[Optional[np.ndarray], Tuple[int, int]]:
    """Captura a imagem do monitor onde está o navegador do sistema."""
    try:
        if x_ref >= 0:
            largura = int(user32.GetSystemMetrics(0)) or 1920
            altura = int(user32.GetSystemMetrics(1)) or 1080
            img_bgr = _capturar_regiao_gdi(0, 0, largura, altura)
            if img_bgr is not None:
                return img_bgr, (0, 0)
            img_pil = ImageGrab.grab()
            return cv2.cvtColor(np.array(img_pil), cv2.COLOR_RGB2BGR), (0, 0)

        vx = int(user32.GetSystemMetrics(76))
        vy = int(user32.GetSystemMetrics(77))
        vw = int(user32.GetSystemMetrics(78))
        vh = int(user32.GetSystemMetrics(79))
        img_bgr = _capturar_regiao_gdi(vx, vy, vw, vh)
        if img_bgr is not None:
            return img_bgr, (vx, vy)
        img_pil = ImageGrab.grab(all_screens=True)
        return cv2.cvtColor(np.array(img_pil), cv2.COLOR_RGB2BGR), (vx, vy)
    except Exception:
        return None, (0, 0)


# Ponto neutro seguro na margem superior branca do cabeçalho do sistema (fora de qualquer input ou link)
PONTO_NEUTRO_PAGINA = (620, 200)


def ir_para_topo_da_pagina(
    x_neutro: Optional[int] = None,
    y_neutro: Optional[int] = None,
) -> None:
    """
    Ativa a janela do navegador clicando em uma área neutra segura (sem inputs/links),
    fecha buscas abertas e garante que a página esteja no topo absoluto (scroll = 0).
    """
    params = obter_parametros_calibrados()
    xn = x_neutro if x_neutro is not None else params["ponto_neutro"][0]
    yn = y_neutro if y_neutro is not None else params["ponto_neutro"][1]
    xn_seguro = max(620, min(1800, int(xn)))
    yn_seguro = max(195, min(350, int(yn)))

    mover_e_clicar(x=xn_seguro, y=yn_seguro, cliques=1, botao="left")
    time.sleep(0.12)
    pressionar_tecla_unica("esc")
    time.sleep(0.08)
    for _ in range(5):
        pressionar_tecla_unica("pageup")
        time.sleep(0.04)
    atalho_teclado("ctrl", "home")
    time.sleep(0.15)
    rolar_pagina(12, x=1230, y=500)
    time.sleep(0.25)
    _log_visao(f"ir_para_topo_da_pagina concluido em ponto neutro ({xn_seguro}, {yn_seguro})")


def rolar_pagina_ate_o_final(
    x_neutro: Optional[int] = None,
    y_neutro: Optional[int] = None,
) -> None:
    """
    Ativa a janela do navegador no ponto neutro seguro, fecha buscas abertas e rola
    a página até o final absoluto (usado após 'Forças' e 'Fraquezas' crescerem a página,
    e ao posicionar o Bloco 2 'Inventário de necessidades').
    """
    params = obter_parametros_calibrados()
    xn = x_neutro if x_neutro is not None else params["ponto_neutro"][0]
    yn = y_neutro if y_neutro is not None else params["ponto_neutro"][1]
    xn_seguro = max(620, min(1800, int(xn)))
    yn_seguro = max(195, min(350, int(yn)))

    mover_e_clicar(x=xn_seguro, y=yn_seguro, cliques=1, botao="left")
    time.sleep(0.12)
    pressionar_tecla_unica("esc")
    time.sleep(0.08)
    for _ in range(5):
        pressionar_tecla_unica("pagedown")
        time.sleep(0.04)
    atalho_teclado("ctrl", "end")
    time.sleep(0.15)
    rolar_pagina(-16, x=1230, y=500)
    time.sleep(0.35)
    _log_visao(f"rolar_pagina_ate_o_final concluido em ponto neutro ({xn_seguro}, {yn_seguro})")


def rolar_planilha_para_direita(
    x_planilha: int,
    y_planilha: int,
    passos_direita: int = 6,
) -> None:
    """
    Rola/navega horizontalmente para a DIREITA dentro da tabela da planilha (ex: '3.2. Inventário de necessidades')
    para alcançar e revelar colunas que ficam mais à direita ('CLASSIFICAÇÃO DE PRIORIDADE',
    'ÁREA DEMANDANTE', 'JUSTIFICATIVA'), sempre pressionando ESC antes para proteger toggles e listas suspensas.
    """
    mover_e_clicar(x=int(x_planilha), y=int(y_planilha), cliques=1, botao="left")
    time.sleep(0.15)
    pressionar_tecla_unica("esc")
    time.sleep(0.08)

    try:
        pyautogui.hscroll(-int(passos_direita) * 120, x=int(x_planilha), y=int(y_planilha))
    except Exception:
        pass

    for _ in range(max(1, int(passos_direita))):
        pressionar_tecla_unica("right")
        time.sleep(0.04)

    time.sleep(0.18)
    _log_visao(
        f"rolar_planilha_para_direita concluido em ({x_planilha}, {y_planilha}) com {passos_direita} passos"
    )


def _aplicar_mascara_topo_navegador(mascara: np.ndarray) -> None:
    """
    Ignora apenas a barra de abas/endereço do navegador (Y < 88) e a caixinha flutuante
    do Ctrl+F no canto superior direito (X >= 1320, Y < 140).
    """
    mascara[:88, :] = 0
    if mascara.shape[1] > 1320:
        mascara[88:140, 1320:] = 0


def detectar_botoes_azuis_adicionar(
    img_bgr: np.ndarray,
    offset_xy: Tuple[int, int],
    incluir_botoes_largos: bool = False,
) -> List[Tuple[int, int, int, int]]:
    """
    Detecta na tela os botões azuis 'Adicionar' / 'Adicionar necessidade' (#2C49F5 -> RGB(44, 73, 245)).
    Por padrão (`incluir_botoes_largos=False`), filtra apenas botões 'Adicionar' das colunas SWOT
    (largura ~75..115px). Quando `incluir_botoes_largos=True`, inclui também botões largos como
    'Adicionar necessidade' (`data-repeater-add="needs"`, largura ~160..225px).
    """
    b = img_bgr[:, :, 0].astype(np.int16)
    g = img_bgr[:, :, 1].astype(np.int16)
    r = img_bgr[:, :, 2].astype(np.int16)

    mascara_azul = (
        (b >= 175)
        & (g >= 20)
        & (g <= 140)
        & (r >= 10)
        & (r <= 115)
        & ((b - r) >= 95)
        & ((b - g) >= 75)
    ).astype(np.uint8) * 255

    _aplicar_mascara_topo_navegador(mascara_azul)

    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (14, 6))
    mascara_fechada = cv2.morphologyEx(mascara_azul, cv2.MORPH_CLOSE, kernel)

    contornos, _ = cv2.findContours(mascara_fechada, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    ox, oy = offset_xy
    w_max = 240 if incluir_botoes_largos else 115
    h_max = 48 if incluir_botoes_largos else 42
    botoes: List[Tuple[int, int, int, int]] = []

    for cnt in contornos:
        x, y, w, h = cv2.boundingRect(cnt)
        if 52 <= w <= w_max and 22 <= h <= h_max:
            recorte = mascara_azul[y : y + h, x : x + w]
            razao_preenchimento = float(np.count_nonzero(recorte)) / float(max(1, w * h))
            if razao_preenchimento < (0.32 if incluir_botoes_largos else 0.45):
                continue
            cx = ox + x + (w // 2)
            cy = oy + y + (h // 2)
            if 600 <= cx <= 1860 and 202 <= cy <= 1065:
                if not incluir_botoes_largos and (cy < 350 and w > 118):
                    continue
                botoes.append((int(cx), int(cy), int(w), int(h)))

    botoes.sort(key=lambda item: item[1])
    return botoes


def detectar_destaques_ctrl_f(
    img_bgr: np.ndarray,
    offset_xy: Tuple[int, int],
    x_min: int = 600,
    x_max: int = 1860,
) -> Tuple[List[Tuple[int, int, int, int]], List[Tuple[int, int, int, int]]]:
    """
    Detecta na página os destaques do Ctrl+F do navegador:
    - Laranja ativo (#FF9632 -> RGB(255, 150, 50))
    - Amarelo inativo (#FFFF00 -> RGB(255, 255, 0))
    Filtra rigorosamente para o monitor do sistema (x_min <= X <= x_max) e ignora sublinhados
    ortográficos pequenos dentro de caixas de texto (que possuem h <= 12).
    """
    b = img_bgr[:, :, 0].astype(np.int16)
    g = img_bgr[:, :, 1].astype(np.int16)
    r = img_bgr[:, :, 2].astype(np.int16)

    mascara_laranja = (
        (r >= 200)
        & (g >= 95)
        & (g <= 205)
        & (b <= 100)
        & ((r - b) >= 120)
    ).astype(np.uint8) * 255

    mascara_amarela = (
        (r >= 200)
        & (g >= 200)
        & (b <= 85)
        & ((r - b) >= 130)
    ).astype(np.uint8) * 255

    _aplicar_mascara_topo_navegador(mascara_laranja)
    _aplicar_mascara_topo_navegador(mascara_amarela)

    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (14, 6))
    ox, oy = offset_xy

    def extrair_retangulos(mascara: np.ndarray) -> List[Tuple[int, int, int, int]]:
        fechada = cv2.morphologyEx(mascara, cv2.MORPH_CLOSE, kernel)
        contornos, _ = cv2.findContours(fechada, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        rects: List[Tuple[int, int, int, int]] = []
        for cnt in contornos:
            x, y, w, h = cv2.boundingRect(cnt)
            # Títulos reais destacados pelo Ctrl+F têm h >= 14 (17px em zoom 90%) e w >= 35px;
            # sublinhados ortográficos vermelhos/laranjas dentro de inputs têm h = 10..11px.
            if 35 <= w <= 420 and 14 <= h <= 50:
                cx = ox + x + (w // 2)
                cy = oy + y + (h // 2)
                if int(x_min) <= cx <= int(x_max) and 85 < cy <= 1065:
                    rects.append((int(cx), int(cy), int(w), int(h)))
        rects.sort(key=lambda item: item[1])
        return rects

    return extrair_retangulos(mascara_laranja), extrair_retangulos(mascara_amarela)


def _executar_ocr_windows_na_regiao(
    img_bgr: np.ndarray,
    offset_xy: Tuple[int, int],
    escala: int = 2,
) -> List[Tuple[int, int, str]]:
    """
    Executa OCR nativo do Windows (`Windows.Media.Ocr.OcrEngine`, idioma `pt-BR`) sobre a imagem capturada
    e retorna uma lista de `(cx_tela, cy_tela, texto_linha)` com as coordenadas de tela exatas de cada linha lida.
    """
    try:
        DADOS_DIR.mkdir(parents=True, exist_ok=True)
        caminho_temp = DADOS_DIR / "temp_ocr_toggle.png"
        fator = max(1, int(escala))
        if fator > 1:
            img_proc = cv2.resize(
                img_bgr,
                (img_bgr.shape[1] * fator, img_bgr.shape[0] * fator),
                interpolation=cv2.INTER_CUBIC,
            )
        else:
            img_proc = img_bgr

        cv2.imwrite(str(caminho_temp), img_proc)
        caminho_ps = str(caminho_temp.resolve()).replace("'", "''")

        ps_script = f"""
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {{ $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' }})[0]
function Await($WinRtTask, $ResultType) {{
    $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
    $netTask = $asTask.Invoke($null, @($WinRtTask))
    $netTask.Wait(-1) | Out-Null
    $netTask.Result
}}
$file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync('{caminho_ps}')) ([Windows.Storage.StorageFile])
$stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
$decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
$bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
$lang = New-Object Windows.Globalization.Language('pt-BR')
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
if ($null -eq $engine) {{ $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages() }}
$result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
foreach ($line in $result.Lines) {{
    $x1 = 99999; $y1 = 99999; $x2 = 0; $y2 = 0
    foreach ($w in $line.Words) {{
        $r = $w.BoundingRect
        if ($r.X -lt $x1) {{ $x1 = $r.X }}
        if ($r.Y -lt $y1) {{ $y1 = $r.Y }}
        if (($r.X + $r.Width) -gt $x2) {{ $x2 = $r.X + $r.Width }}
        if (($r.Y + $r.Height) -gt $y2) {{ $y2 = $r.Y + $r.Height }}
    }}
    $cx = [int](($x1 + $x2) / 2)
    $cy = [int](($y1 + $y2) / 2)
    Write-Output ("{{0}}|{{1}}|{{2}}" -f $cx, $cy, $line.Text)
}}
$stream.Dispose()
"""
        encoded = base64.b64encode(ps_script.encode("utf-16le")).decode("ascii")
        proc = subprocess.run(
            ["powershell", "-NoProfile", "-EncodedCommand", encoded],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=8,
        )
        ox, oy = offset_xy
        linhas_ocr: List[Tuple[int, int, str]] = []
        for raw_line in (proc.stdout or "").splitlines():
            partes = raw_line.strip().split("|", 2)
            if len(partes) == 3:
                try:
                    cx_img = int(float(partes[0])) // fator
                    cy_img = int(float(partes[1])) // fator
                    txt = partes[2].strip()
                    if txt:
                        linhas_ocr.append((ox + cx_img, oy + cy_img, txt))
                except ValueError:
                    continue
        return linhas_ocr
    except Exception as exc:
        _log_visao(f"Erro em _executar_ocr_windows_na_regiao: {exc}")
        return []


def _corresponde_nome_guia_ocr(texto_ocr: str, nome_guia_referencia: str) -> Tuple[bool, int]:
    """
    Avalia se uma linha lida pelo OCR dentro do menu Toggle de Todas as Guias corresponde
    ao `nome_guia_referencia` (ex: '3.1. Análise SWOT', '3.1. SWOT', '3.2. Inventário de necessidades',
    '4.1. Metas e Ações'). Retorna `(encontrou, pontuacao)`.
    """
    t_norm = _norm_chave(texto_ocr)
    ref_norm = _norm_chave(nome_guia_referencia)
    if not t_norm or not ref_norm:
        return False, 0

    if "BANCO" in t_norm and "BANCO" not in ref_norm:
        return False, 0

    if t_norm == ref_norm:
        return True, 100
    if "4.1" in ref_norm and "METAS" in ref_norm and "4.1" in t_norm and "METAS E A" in t_norm:
        return True, 98
    if ref_norm in t_norm or (len(t_norm) >= 5 and t_norm in ref_norm):
        return True, 90

    # Extrai prefixo numérico (ex: '3.1', '3.2', '4.1', '5.1', '6.1')
    m_ref = re.search(r"\b(\d+\.\d+)\b", ref_norm)
    prefixo_num = m_ref.group(1) if m_ref else ""

    palavras_chave = [
        p
        for p in re.split(r"[^A-Z0-9]+", ref_norm)
        if len(p) >= 4 and p not in ("ANALISE", "PARA", "COMO", "PELO", "PELA")
    ]

    tem_prefixo = bool(prefixo_num and prefixo_num in t_norm)
    qtd_palavras = sum(1 for p in palavras_chave if p in t_norm)

    if tem_prefixo and qtd_palavras >= 1:
        return True, 85 + qtd_palavras * 5
    if qtd_palavras >= 2:
        return True, 75 + qtd_palavras * 5
    if "SWOT" in ref_norm and "SWOT" in t_norm:
        return True, 88
    if "INVENTARIO" in ref_norm and ("INVENTARIO" in t_norm or "NECESSIDADE" in t_norm):
        return True, 85
    if "METAS" in ref_norm and ("METAS" in t_norm or "ACOES" in t_norm):
        return True, 85
    if tem_prefixo and not palavras_chave:
        return True, 70

    return False, 0


def selecionar_guia_planilha_via_toggle_e_nome(
    nome_guia: str,
    coord_toggle_todas_guias: Optional[Tuple[int, int]] = None,
    coord_item_no_toggle: Optional[Tuple[int, int]] = None,
    coord_guia_fallback: Optional[Tuple[int, int]] = None,
) -> Optional[Tuple[int, int]]:
    """
    Seleciona a guia/aba da planilha SEM depender da posição variável na barra horizontal inferior:
    1. Clica no botão fixo 'Toggle de Todas as Guias' (`≡` — `coord_toggle_todas_guias`) na planilha,
       que abre o menu vertical listando todas as abas.
    2. Captura a imagem do menu aberto e usa OCR nativo do Windows (`pt-BR`) para localizar
       o nome de referência `nome_guia` (ex: '3.1. Análise SWOT' / '3.1. SWOT',
       '3.2. Inventário de necessidades', '4.1. Metas e Ações').
    3. Clica exatamente sobre o nome da guia encontrada no menu do Toggle!
    4. Caso o item específico dentro do Toggle tenha sido ensinado (`coord_item_no_toggle`) ou
       haja `coord_guia_fallback`, utiliza como fallback seguro.
    """
    if (
        isinstance(coord_toggle_todas_guias, (list, tuple))
        and len(coord_toggle_todas_guias) >= 2
        and coord_toggle_todas_guias[0] is not None
        and coord_toggle_todas_guias[1] is not None
    ):
        tx, ty = int(coord_toggle_todas_guias[0]), int(coord_toggle_todas_guias[1])
        print(
            f"📑 [TOGGLE DE TODAS AS GUIAS] Clicando no Toggle '≡' da planilha em ({tx}, {ty}) "
            f"para buscar a guia pelo nome '{nome_guia}'..."
        )
        mover_e_clicar(x=tx, y=ty, cliques=1, botao="left")
        time.sleep(0.55)

        # Captura a região onde o menu suspenso de todas as guias se abre (acima/ao lado do botão Toggle)
        x1 = max(-1920, tx - 70)
        y1 = max(70, ty - 780)
        x2 = min(1920, tx + 540)
        y2 = min(1075, ty + 15)
        w_reg = max(120, x2 - x1)
        h_reg = max(120, y2 - y1)

        melhor_coord: Optional[Tuple[int, int]] = None
        melhor_txt = ""
        melhor_score = 0

        for tentativa_scroll in range(2):
            img_menu = _capturar_regiao_gdi(x1, y1, w_reg, h_reg)
            if img_menu is not None:
                linhas_ocr = _executar_ocr_windows_na_regiao(img_menu, (x1, y1), escala=2)
                if linhas_ocr:
                    textos_vistos = [t for _, _, t in linhas_ocr]
                    _log_visao(f"OCR Toggle todas as guias (tentativa {tentativa_scroll + 1}): {textos_vistos}")
                for cx_item, cy_item, txt_item in linhas_ocr:
                    # Evita clicar novamente no próprio botão de toggle na borda inferior
                    if abs(cy_item - ty) <= 12 and abs(cx_item - tx) <= 25:
                        continue
                    ok_match, score = _corresponde_nome_guia_ocr(txt_item, nome_guia)
                    if ok_match and score > melhor_score:
                        melhor_score = score
                        melhor_coord = (int(cx_item), int(cy_item))
                        melhor_txt = txt_item

            if melhor_coord is not None:
                break

            # Se a lista do toggle tiver barra de rolagem, rola suavemente o menu na 2ª tentativa
            if tentativa_scroll == 0:
                rolar_pagina(-5, x=tx + 120, y=max(200, ty - 220))
                time.sleep(0.35)

        if melhor_coord is not None:
            print(
                f"   ✅ Guia '{nome_guia}' localizada pelo nome no menu do Toggle ('{melhor_txt}') "
                f"em {melhor_coord}! Clicando..."
            )
            mover_e_clicar(x=melhor_coord[0], y=melhor_coord[1], cliques=1, botao="left")
            time.sleep(0.65)
            _log_visao(
                f"selecionar_guia_planilha_via_toggle_e_nome '{nome_guia}' -> '{melhor_txt}' em {melhor_coord}"
            )
            return melhor_coord

        if (
            isinstance(coord_item_no_toggle, (list, tuple))
            and len(coord_item_no_toggle) >= 2
            and coord_item_no_toggle[0] is not None
            and coord_item_no_toggle[1] is not None
        ):
            ix, iy = int(coord_item_no_toggle[0]), int(coord_item_no_toggle[1])
            print(
                f"   👉 Clicando na posição calibrada do item '{nome_guia}' dentro do menu Toggle em ({ix}, {iy})..."
            )
            mover_e_clicar(x=ix, y=iy, cliques=1, botao="left")
            time.sleep(0.65)
            return (ix, iy)

        # Fecha o menu do toggle com ESC caso vá usar o fallback da barra de abas
        pressionar_tecla_unica("esc")
        time.sleep(0.18)

    if (
        isinstance(coord_guia_fallback, (list, tuple))
        and len(coord_guia_fallback) >= 2
        and coord_guia_fallback[0] is not None
        and coord_guia_fallback[1] is not None
    ):
        gx, gy = int(coord_guia_fallback[0]), int(coord_guia_fallback[1])
        print(f"🔗 Clicando na guia '{nome_guia}' na planilha em ({gx}, {gy})...")
        mover_e_clicar(x=gx, y=gy, cliques=1, botao="left")
        time.sleep(0.65)
        return (gx, gy)

    return None


def _confirmar_coluna_ou_secao_na_guia_sistema(titulo_coluna: str) -> bool:
    """
    Confirma visualmente via Ctrl+F + OpenCV que a página/guia aberta no sistema
    contém o cabeçalho da seção ou coluna esperada (ex: 'Forças' na guia '3. Diagnóstico'
    ou 'Metas e ações' na guia '4. Metas e Planos') dentro da área central do formulário (X >= 620).
    """
    termo = str(titulo_coluna or "").strip()
    if not termo:
        return True

    copiar_para_clipboard(termo)
    time.sleep(0.05)
    atalho_teclado("ctrl", "f")
    time.sleep(0.16)
    atalho_teclado("ctrl", "a")
    time.sleep(0.04)
    atalho_teclado("ctrl", "v")
    time.sleep(0.28)

    encontrado = False
    img_bgr, off_xy = capturar_tela_sistema(620)
    if img_bgr is not None:
        laranjas, amarelos = detectar_destaques_ctrl_f(
            img_bgr, off_xy, x_min=620, x_max=1860
        )
        candidatos = [
            r for r in (laranjas + amarelos) if 620 <= r[0] <= 1860 and 140 <= r[1] <= 1040
        ]
        encontrado = bool(candidatos)

    pressionar_tecla_unica("esc")
    time.sleep(0.12)
    return encontrado


def selecionar_guia_wizard_no_sistema(
    nome_guia_sistema: str,
    coord_guia_sistema: Optional[Tuple[int, int]] = None,
    step_index: Optional[int] = None,
    titulo_coluna_confirmacao: Optional[str] = None,
    forcar_clique: bool = False,
) -> Tuple[int, int]:
    """
    Verifica PRIMEIRO se o sistema já está na guia certa e, caso não esteja, seleciona e confirma
    a guia/etapa correspondente no menu lateral do sistema (`<div class="nav flex-column wizard-nav">`):
    1. Garante que a janela do navegador está ativa e no topo (`ir_para_topo_da_pagina`).
    2. Se `titulo_coluna_confirmacao` foi informado e `forcar_clique=False`, verifica primeiro se a
       seção/coluna esperada já está aberta na tela. Se já estiver, não precisa clicar novamente!
    3. Caso não esteja na guia certa, clica em `coord_guia_sistema` (se calibrada) ou localiza o botão
       em `.wizard-nav` via Ctrl+F + OpenCV (`3. Diagnóstico`, `4. Metas e Planos`,
       `5. Capacitação e Comunicação`, `6. Monitoramento e Revisão`) e confirma a abertura.
    """
    ir_para_topo_da_pagina()

    # 1º Passo no sistema: verificar se já está na guia certa antes de clicar
    if not forcar_clique and titulo_coluna_confirmacao:
        print(
            f"🔎 [VERIFICAÇÃO DO SISTEMA] Verificando se o sistema já está na guia '{nome_guia_sistema}' "
            f"(seção '{titulo_coluna_confirmacao}')..."
        )
        if _confirmar_coluna_ou_secao_na_guia_sistema(titulo_coluna_confirmacao):
            print(
                f"   ✅ O sistema JÁ está na guia certa ('{nome_guia_sistema}') "
                f"com '{titulo_coluna_confirmacao}' visível na tela!"
            )
            if (
                isinstance(coord_guia_sistema, (list, tuple))
                and len(coord_guia_sistema) >= 2
                and coord_guia_sistema[0] is not None
                and coord_guia_sistema[1] is not None
            ):
                return (int(coord_guia_sistema[0]), int(coord_guia_sistema[1]))
            idx_etapa_atual = int(step_index) if step_index is not None else 2
            return (435, 330 + idx_etapa_atual * 48)

    if (
        isinstance(coord_guia_sistema, (list, tuple))
        and len(coord_guia_sistema) >= 2
        and coord_guia_sistema[0] is not None
        and coord_guia_sistema[1] is not None
    ):
        gx, gy = int(coord_guia_sistema[0]), int(coord_guia_sistema[1])
        print(
            f"🖥️  [GUIA DO SISTEMA] Clicando na etapa '{nome_guia_sistema}' (wizard-nav) em ({gx}, {gy})..."
        )
        mover_e_clicar(x=gx, y=gy, cliques=1, botao="left")
        time.sleep(0.50)
        _log_visao(f"selecionar_guia_wizard_no_sistema (calibrado) '{nome_guia_sistema}' em ({gx}, {gy})")
        if not titulo_coluna_confirmacao or _confirmar_coluna_ou_secao_na_guia_sistema(titulo_coluna_confirmacao):
            if titulo_coluna_confirmacao:
                print(
                    f"   ✅ Confirmado no sistema: guia '{nome_guia_sistema}' ativa e coluna/seção "
                    f"'{titulo_coluna_confirmacao}' visível na tela!"
                )
            return (gx, gy)
        print(
            f"   ⚠️ Seção/coluna '{titulo_coluna_confirmacao}' não detectada após clique em ({gx}, {gy}). "
            f"Buscando botão '{nome_guia_sistema}' em '.wizard-nav' via Ctrl+F + OpenCV..."
        )

    termo_busca = str(nome_guia_sistema or "").strip()
    if "METAS" in _norm_chave(termo_busca):
        termo_busca = "4. Metas e Planos"
        if step_index is None:
            step_index = 3
    elif "DIAGNOSTICO" in _norm_chave(termo_busca):
        termo_busca = "3. Diagnóstico"
        if step_index is None:
            step_index = 2
    elif "FUNDAMENTACAO" in _norm_chave(termo_busca):
        termo_busca = "2. Fundamentação"
        if step_index is None:
            step_index = 1
    elif "IDENTIFICACAO" in _norm_chave(termo_busca):
        termo_busca = "1. Identificação"
        if step_index is None:
            step_index = 0
    elif "CAPACITACAO" in _norm_chave(termo_busca):
        termo_busca = "5. Capacitação e Comunicação"
        if step_index is None:
            step_index = 4
    elif "MONITORAMENTO" in _norm_chave(termo_busca):
        termo_busca = "6. Monitoramento e Revisão"
        if step_index is None:
            step_index = 5

    print(
        f"🔎 [GUIA DO SISTEMA] Localizando botão '{termo_busca}' na div '.wizard-nav' para abrir os campos..."
    )
    copiar_para_clipboard(termo_busca)
    time.sleep(0.06)
    atalho_teclado("ctrl", "f")
    time.sleep(0.18)
    atalho_teclado("ctrl", "a")
    time.sleep(0.04)
    atalho_teclado("ctrl", "v")
    time.sleep(0.32)

    coord_encontrada: Optional[Tuple[int, int]] = None
    for tentativa in range(3):
        img_bgr, off_xy = capturar_tela_sistema(620)
        if img_bgr is not None:
            laranjas, amarelos = detectar_destaques_ctrl_f(
                img_bgr, off_xy, x_min=240, x_max=615
            )
            candidatos = [
                r for r in (laranjas + amarelos) if 240 <= r[0] <= 615 and 200 <= r[1] <= 760
            ]
            if candidatos:
                candidatos.sort(key=lambda r: r[1])
                coord_encontrada = (int(candidatos[0][0]), int(candidatos[0][1]))
                break
        if tentativa < 2:
            pressionar_tecla_unica("enter")
            time.sleep(0.22)

    pressionar_tecla_unica("esc")
    time.sleep(0.15)

    if coord_encontrada is None:
        idx_etapa = int(step_index) if step_index is not None else 3
        coord_encontrada = (435, 330 + idx_etapa * 48)
        print(
            f"   ℹ️  Usando posição padrão da etapa {idx_etapa + 1} no '.wizard-nav': {coord_encontrada}"
        )
    else:
        print(
            f"   ✅ Botão '{termo_busca}' localizado em '.wizard-nav' na coordenada {coord_encontrada}!"
        )

    mover_e_clicar(x=coord_encontrada[0], y=coord_encontrada[1], cliques=1, botao="left")
    time.sleep(0.50)
    if titulo_coluna_confirmacao:
        if _confirmar_coluna_ou_secao_na_guia_sistema(titulo_coluna_confirmacao):
            print(
                f"   ✅ Confirmado no sistema: guia '{termo_busca}' ativa e coluna/seção "
                f"'{titulo_coluna_confirmacao}' visível na tela!"
            )
    _log_visao(
        f"selecionar_guia_wizard_no_sistema (auto) '{termo_busca}' clicado em {coord_encontrada}"
    )
    return coord_encontrada


def clicar_botao_proxima_etapa_no_sistema(
    coord_botao_salva: Optional[Tuple[int, int]] = None,
) -> Optional[Tuple[int, int]]:
    """
    Avança para a próxima fase/etapa do sistema clicando no botão ao final da página:
    `<button type="button" class="btn btn-primary" data-step-action="next">Próxima etapa</button>`.
    1. Rola a página até o final (`rolar_pagina_ate_o_final`) onde fica o rodapé de navegação.
    2. Se `coord_botao_salva` foi calibrada, clica nela; caso contrário, localiza 'Próxima etapa'
       via Ctrl+F + OpenCV na área do formulário (X >= 620) e clica no botão.
    """
    rolar_pagina_ate_o_final()

    if (
        isinstance(coord_botao_salva, (list, tuple))
        and len(coord_botao_salva) >= 2
        and coord_botao_salva[0] is not None
        and coord_botao_salva[1] is not None
    ):
        bx, by = int(coord_botao_salva[0]), int(coord_botao_salva[1])
        print(
            f"⏭️  [PRÓXIMA ETAPA] Clicando no botão 'Próxima etapa' (`data-step-action=\"next\"`) em ({bx}, {by})..."
        )
        mover_e_clicar(x=bx, y=by, cliques=1, botao="left")
        time.sleep(0.55)
        ir_para_topo_da_pagina()
        return (bx, by)

    print(
        "🔎 [PRÓXIMA ETAPA] Localizando botão '<button data-step-action=\"next\">Próxima etapa</button>'..."
    )
    copiar_para_clipboard("Próxima etapa")
    time.sleep(0.05)
    atalho_teclado("ctrl", "f")
    time.sleep(0.16)
    atalho_teclado("ctrl", "a")
    time.sleep(0.04)
    atalho_teclado("ctrl", "v")
    time.sleep(0.30)

    coord_btn: Optional[Tuple[int, int]] = None
    img_bgr, off_xy = capturar_tela_sistema(620)
    if img_bgr is not None:
        laranjas, amarelos = detectar_destaques_ctrl_f(img_bgr, off_xy, x_min=620, x_max=1890)
        candidatos = [r for r in (laranjas + amarelos) if r[0] >= 620 and r[1] >= 300]
        if candidatos:
            candidatos.sort(key=lambda r: (r[1], r[0]), reverse=True)
            coord_btn = (int(candidatos[0][0]), int(candidatos[0][1]))

    pressionar_tecla_unica("esc")
    time.sleep(0.12)

    if coord_btn is not None:
        print(f"   ✅ Botão 'Próxima etapa' localizado em {coord_btn}! Clicando para avançar...")
        mover_e_clicar(x=coord_btn[0], y=coord_btn[1], cliques=1, botao="left")
        time.sleep(0.55)
        ir_para_topo_da_pagina()
        _log_visao(f"clicar_botao_proxima_etapa_no_sistema clicado em {coord_btn}")
        return coord_btn

    return None


def detectar_botoes_remover_da_coluna(
    img_bgr: np.ndarray,
    offset_xy: Tuple[int, int],
    x_btn_add: int,
    y_btn_add: int,
    y_limite_inferior: int = 1010,
    passo_max_continuo: Optional[int] = 115,
) -> List[int]:
    """
    Detecta a coordenada Y de cada botão '×' vermelho (#e0314f / borda #f2b7c1) alinhado abaixo
    do botão 'Adicionar' daquele card específico (limitado até y_limite_inferior).
    Se `passo_max_continuo` for informado (ex: 115px para colunas individuais), interrompe a lista
    caso haja um salto vertical maior que `passo_max_continuo` (evitando capturar '×' de outra seção abaixo).
    """
    b = img_bgr[:, :, 0].astype(np.int16)
    g = img_bgr[:, :, 1].astype(np.int16)
    r = img_bgr[:, :, 2].astype(np.int16)

    # Captura tanto o ícone '×' (#e0314f: R=224,G=49,B=79) quanto a borda rosa (#f2b7c1: R=242,G=183,B=193)
    mascara_vermelha = (
        (r >= 170)
        & (g <= 230)
        & (b <= 235)
        & ((r - g) >= 16)
        & ((r - b) >= 12)
    ).astype(np.uint8) * 255

    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (10, 10))
    fechada = cv2.morphologyEx(mascara_vermelha, cv2.MORPH_CLOSE, kernel)
    contornos, _ = cv2.findContours(fechada, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    ox, oy = offset_xy
    limite_y = min(1010, int(y_limite_inferior))
    y_min_busca = max(95, int(y_btn_add) + 40)
    ys_encontrados: List[int] = []
    for cnt in contornos:
        x, y, w, h = cv2.boundingRect(cnt)
        if 1 <= w <= 68 and 1 <= h <= 68:
            cx = ox + x + (w // 2)
            cy = oy + y + (h // 2)
            # Nos cards largos (x_btn_add ~ 1769), o botão '×' de 11 colunas (Metas e Ações / Plano de Riscos)
            # fica na coluna 11 (cx ~ 1672, distância ~ 97px), enquanto de 12 colunas fica na coluna 12 (cx ~ 1775).
            tol_x = 148 if int(x_btn_add) >= 1500 else 105
            if abs(cx - int(x_btn_add)) <= tol_x and y_min_busca <= cy <= limite_y:
                ys_encontrados.append(int(cy))

    ys_encontrados.sort()
    ys_agrupados: List[int] = []
    for y_val in ys_encontrados:
        if not ys_agrupados or abs(y_val - ys_agrupados[-1]) > 35:
            ys_agrupados.append(y_val)

    if passo_max_continuo is not None and len(ys_agrupados) >= 2:
        ys_continuos: List[int] = [ys_agrupados[0]]
        for y_val in ys_agrupados[1:]:
            if (y_val - ys_continuos[-1]) <= passo_max_continuo:
                ys_continuos.append(y_val)
            else:
                break
        return ys_continuos

    return ys_agrupados


def centralizar_y_na_caixa_branca(
    img_bgr: np.ndarray,
    offset_xy: Tuple[int, int],
    x_caixa: int,
    y_caixa: int,
) -> int:
    """
    Ajusta verticalmente `y_caixa` para o centro exato da caixa
    de texto branca (<input class="form-control">) sob `x_caixa`:
    1. Procura primeiro o par de bordas horizontais superior/inferior do <input> (separadas por ~36..44px)
       com miolo branco contínuo entre elas, retornando o centro geométrico exato da caixa.
    2. Como fallback, busca a faixa inteiramente branca (incluindo o centro yc±6, sem borda no meio!)
       mais próxima de `y_caixa`.
    """
    ox, oy = offset_xy
    h_img, w_img = img_bgr.shape[:2]
    x1 = max(0, int(x_caixa) - ox - 35)
    x2 = min(w_img, int(x_caixa) - ox + 55)
    if (x2 - x1) < 25:
        return int(y_caixa)

    faixa = img_bgr[:, x1:x2]
    b_f = faixa[:, :, 0].astype(np.int16)
    g_f = faixa[:, :, 1].astype(np.int16)
    r_f = faixa[:, :, 2].astype(np.int16)
    gray_f = cv2.cvtColor(faixa, cv2.COLOR_BGR2GRAY)

    eh_borda = (
        ((r_f >= 190) & (r_f <= 240) & (g_f >= 200) & (g_f <= 245) & (b_f >= 215) & (b_f <= 252) & ((b_f - r_f) >= 5))
        | ((b_f >= 185) & ((b_f - r_f) >= 65))
    )
    score_borda = np.mean(eh_borda, axis=1)
    score_branco = np.mean(gray_f >= 243, axis=1)

    melhor_centro_borda: Optional[int] = None
    menor_dist_borda = 999999.0
    for dy in range(-28, 29):
        yc_abs = int(y_caixa) + dy
        yc_rel = yc_abs - oy
        if yc_rel - 23 < 0 or yc_rel + 23 >= h_img:
            continue
        borda_sup = float(np.max(score_borda[yc_rel - 22 : yc_rel - 17]))
        borda_inf = float(np.max(score_borda[yc_rel + 17 : yc_rel + 23]))
        borda_meio = float(np.max(score_borda[yc_rel - 13 : yc_rel + 14]))
        miolo_branco = float(np.mean(score_branco[yc_rel - 13 : yc_rel + 14]))
        if borda_sup >= 0.42 and borda_inf >= 0.42 and borda_meio < 0.35 and miolo_branco >= 0.72:
            dist = abs(dy) - (borda_sup + borda_inf) * 2.0
            if dist < menor_dist_borda:
                menor_dist_borda = dist
                melhor_centro_borda = yc_abs

    if melhor_centro_borda is not None:
        return int(melhor_centro_borda)

    melhor_y = int(y_caixa)
    melhor_score = -1.0
    for dy in range(-18, 19, 2):
        yc_abs = int(y_caixa) + dy
        yc_rel = yc_abs - oy
        if yc_rel - 13 < 0 or yc_rel + 13 >= h_img:
            continue
        branco_sup = float(np.mean(score_branco[yc_rel - 13 : yc_rel - 6]))
        branco_meio = float(np.mean(score_branco[yc_rel - 6 : yc_rel + 7]))
        branco_inf = float(np.mean(score_branco[yc_rel + 6 : yc_rel + 13]))
        borda_meio = float(np.max(score_borda[yc_rel - 10 : yc_rel + 11]))
        if borda_meio >= 0.45:
            continue
        penalidade_dist = abs(dy) * 0.004
        score = min(branco_sup, branco_meio, branco_inf) * 0.7 + ((branco_sup + branco_inf) * 0.15) - penalidade_dist
        if score > melhor_score and branco_sup >= 0.75 and branco_meio >= 0.70 and branco_inf >= 0.75:
            melhor_score = score
            melhor_y = yc_abs

    return int(melhor_y)


def eh_caixa_de_texto_vazia(
    img_bgr: np.ndarray,
    offset_xy: Tuple[int, int],
    x_caixa: int,
    y_caixa: int,
    x_btn_add: Optional[int] = None,
) -> bool:
    """
    Verifica visualmente se a coordenada `(x_caixa, y_caixa)` aponta para o interior de uma
    CAIXA DE TEXTO VAZIA (<input class="form-control"> com fundo #ffffff e sem texto escuro #03183C).
    Detecta automaticamente as bordas esquerda/direita da caixa ao redor de `x_caixa` para funcionar
    tanto em colunas largas (SWOT) quanto em blocos com vários campos lado a lado (Página 4: ID, Prazo, etc.).
    """
    ox, oy = offset_xy
    h_img, w_img = img_bgr.shape[:2]

    y_centro = centralizar_y_na_caixa_branca(img_bgr, offset_xy, x_caixa, y_caixa)
    yc_rel = int(y_centro) - oy
    xc_rel = int(x_caixa) - ox
    if yc_rel - 12 < 0 or yc_rel + 12 >= h_img or xc_rel < 10 or xc_rel >= (w_img - 10):
        return False

    # Delimita as bordas horizontais [x_esq_rel, x_dir_rel] da caixa branca que contém `x_caixa`
    # usando o percentil 85 vertical (que permanece >= 242 no padding branco superior/inferior mesmo se houver texto,
    # mas cai < 238 na linha vertical da borda esquerda/direita do <input>).
    # Usa alcance esquerdo de até 540px para alcançar a borda esquerda real (X ~ 675) mesmo em caixas largas col-md-12/col-md-8.
    x_busca_min = max(max(0, 650 - ox), xc_rel - 540)
    x_busca_max = min(w_img - 10, xc_rel + 280)
    if x_btn_add is not None and int(x_btn_add) > int(x_caixa) + 60:
        x_busca_max = min(x_busca_max, max(xc_rel + 25, int(x_btn_add) - ox - 68))

    faixa_input = cv2.cvtColor(
        img_bgr[yc_rel - 12 : yc_rel + 12, x_busca_min:x_busca_max],
        cv2.COLOR_BGR2GRAY,
    )
    if faixa_input.size == 0:
        return False

    perfil_branco_x = np.percentile(faixa_input, 85, axis=0)
    idx_xc = max(0, min(len(perfil_branco_x) - 1, xc_rel - x_busca_min))

    # Se nem o próprio ponto (x_caixa, y_centro) tem interior branco, não é uma caixa de texto branca
    if float(np.mean(perfil_branco_x[max(0, idx_xc - 8) : min(len(perfil_branco_x), idx_xc + 9)])) < 240.0:
        return False

    idx_esq = idx_xc
    while idx_esq > 0 and perfil_branco_x[idx_esq - 1] >= 240:
        idx_esq -= 1

    idx_dir = idx_xc
    while idx_dir < (len(perfil_branco_x) - 1) and perfil_branco_x[idx_dir + 1] >= 240:
        idx_dir += 1

    # Recorta o miolo interno da caixa começando logo após a borda esquerda (onde qualquer texto digitado começa!)
    x1 = x_busca_min + idx_esq + 7
    x2 = min(x_busca_min + idx_dir - 7, x1 + 290)
    if (x2 - x1) < 28:
        x1 = max(0, xc_rel - 45)
        x2 = min(w_img, xc_rel + 85)

    y1 = max(0, yc_rel - 9)
    y2 = min(h_img, yc_rel + 9)
    if (x2 - x1) < 24 or (y2 - y1) < 8:
        return False

    recorte = img_bgr[y1:y2, x1:x2]
    gray = cv2.cvtColor(recorte, cv2.COLOR_BGR2GRAY)

    # Fundo branco do input (#ffffff -> gray >= 242) e ausência de caracteres digitados reais (#212529 / #03183C -> gray < 78 / < 94).
    # Nota: placeholders do Bootstrap 5 (ex: placeholder="A1", "0,00", "Ex.: Tesouro / Convênio") têm cor #6c757d (gray ~ 108..125)
    # e NÃO possuem pixels < 78 nem < 94, garantindo que caixas vazias com placeholder sejam reconhecidas como VAZIAS!
    proporcao_branco = float(np.mean(gray >= 242))
    pixels_texto_forte = int(np.sum(gray < 78))
    pixels_texto_medio = int(np.sum(gray < 94))
    return proporcao_branco >= 0.82 and pixels_texto_forte < 6 and pixels_texto_medio < 14


def detectar_caixas_de_texto_da_coluna(
    img_bgr: np.ndarray,
    offset_xy: Tuple[int, int],
    x_guia_caixa: int,
    x_btn_add: int,
    y_btn_add: int,
    y_limite_inferior: int = 1010,
    apenas_vazias: bool = False,
) -> List[int]:
    """
    Localiza visualmente as coordenadas Y das caixas de texto (<input class="form-control">)
    da coluna atual:
    1. Prioriza exclusivamente os botões '×' (#f2b7c1 / #e0314f) alinhados à direita de cada caixa da coluna
       (evitando confundir o espaço branco de 34px entre caixas ou invadir caixas/textareas de outra seção abaixo,
       como 'Inventário de necessidades').
    2. Apenas se a coluna não tiver botões '×' (campo simples sem repetidor), usa a varredura direta das faixas brancas.
    Se `apenas_vazias=True`, retorna somente as caixas de texto que estão vazias (prontas para colar).
    """
    ys_candidatos = list(
        detectar_botoes_remover_da_coluna(
            img_bgr, offset_xy, x_btn_add, y_btn_add, y_limite_inferior
        )
    )

    # Varredura de faixa apenas como fallback para campos únicos que não possuem botão '×'
    if not ys_candidatos:
        ox, oy = offset_xy
        h_img, w_img = img_bgr.shape[:2]
        x1 = max(0, int(x_guia_caixa) - ox - 110)
        x2 = min(w_img, int(x_guia_caixa) - ox + 110)
        y_min = max(95, int(y_btn_add) + 42)
        y_max = min(1005, int(y_limite_inferior))

        if (x2 - x1) >= 60 and y_max > (y_min + 25):
            faixa = img_bgr[:, x1:x2]
            b_f = faixa[:, :, 0].astype(np.int16)
            g_f = faixa[:, :, 1].astype(np.int16)
            r_f = faixa[:, :, 2].astype(np.int16)

            eh_borda = (
                ((r_f >= 200) & (r_f <= 236) & (g_f >= 210) & (g_f <= 242) & (b_f >= 224) & (b_f <= 250) & ((b_f - r_f) >= 8))
                | ((b_f >= 190) & ((b_f - r_f) >= 80))
            )
            score_borda = np.mean(eh_borda, axis=1)
            eh_branco = (r_f >= 248) & (g_f >= 248) & (b_f >= 248)
            score_branco = np.mean(eh_branco, axis=1)

            for cy_abs in range(y_min, y_max - 20):
                cy_rel = cy_abs - oy
                if cy_rel - 21 < 0 or cy_rel + 21 >= h_img:
                    continue
                borda_sup = float(np.max(score_borda[cy_rel - 22 : cy_rel - 19]))
                borda_inf = float(np.max(score_borda[cy_rel + 19 : cy_rel + 23]))
                miolo_branco = float(np.mean(score_branco[cy_rel - 12 : cy_rel + 12]))
                if borda_sup >= 0.45 and borda_inf >= 0.45 and miolo_branco >= 0.70:
                    if all(abs(cy_abs - y_ex) > 42 for y_ex in ys_candidatos):
                        ys_candidatos.append(int(cy_abs))

    ys_candidatos.sort()
    ys_filtrados: List[int] = []
    for y_box in ys_candidatos:
        if ys_filtrados and abs(y_box - ys_filtrados[-1]) <= 45:
            continue
        if apenas_vazias:
            if eh_caixa_de_texto_vazia(img_bgr, offset_xy, x_guia_caixa, y_box, x_btn_add):
                ys_filtrados.append(int(y_box))
        else:
            ys_filtrados.append(int(y_box))

    return ys_filtrados


def inspecionar_caixa_vazia_no_ponto(cx: int, cy: int) -> Dict[str, Any]:
    """
    Usado pelo mapeador (`treinador.py`) quando o usuário aponta para uma caixa de texto vazia inicial:
    Inspeciona a imagem ao redor de `(cx, cy)` para confirmar a caixa vazia e detectar automaticamente
    o botão azul 'Adicionar' correspondente a essa coluna.
    """
    img, offset_xy = capturar_tela_sistema(cx)
    resultado: Dict[str, Any] = {
        "campo": [int(cx), int(cy)],
        "botao_adicionar": None,
        "caixa_vazia_confirmada": False,
    }
    if img is None:
        return resultado

    botoes = detectar_botoes_azuis_adicionar(img, offset_xy, incluir_botoes_largos=False)
    # O botão 'Adicionar' do card fica à direita do meio da caixa (bx > cx + 80) e ~50..115px acima da 1ª caixa
    candidatos_btn = [
        b
        for b in botoes
        if (cx + 60) <= b[0] <= (cx + 420) and 35 <= (cy - b[1]) <= 135
    ]
    if candidatos_btn:
        candidatos_btn.sort(key=lambda b: abs((cy - b[1]) - 73) + abs((b[0] - cx) - 256))
        melhor_b = candidatos_btn[0]
        resultado["botao_adicionar"] = [int(melhor_b[0]), int(melhor_b[1])]
        x_btn_ref = int(melhor_b[0])
    else:
        eh_col_direita = cx >= 1240
        x_btn_ref = 1775 if eh_col_direita else 1174
        resultado["botao_adicionar"] = [x_btn_ref, int(cy - 73)]

    resultado["caixa_vazia_confirmada"] = eh_caixa_de_texto_vazia(
        img, offset_xy, cx, cy, x_btn_ref
    )
    return resultado


def _obter_limite_y_proximo_card(
    todos_botoes: List[Tuple[int, int, int, int]],
    x_btn_add: int,
    y_btn_add: int,
    y_limite_extra: Optional[int] = None,
    y_minimo_abaixo: Optional[int] = None,
) -> int:
    """
    Retorna o Y máximo permitido acima do próximo card/bloco vizinho abaixo:
    - Detecta qualquer botão azul na mesma coluna (`abs(b[0] - x_btn_add) <= 180`) OU
      qualquer botão azul largo de bloco de página inteira (`b[2] >= 112`, ex: 'Adicionar necessidade',
      'Adicionar linha' do Plano Orçamentário, 'Adicionar risco' da Gestão de Riscos) que esteja abaixo de `y_btn_add`.
    - Ignora botões no cabeçalho superior (`b[1] < 260`) para nunca confundir o próprio botão do bloco
      (quando próximo ao topo após rolagem) com um bloco vizinho abaixo.
    - Respeita também `y_limite_extra` (barreira de bloco vizinho ensinada no treinamento).
    """
    y_piso = max(260, int(y_btn_add) + 92)
    if y_minimo_abaixo is not None:
        y_piso = max(y_piso, int(y_minimo_abaixo))

    proximos = [
        b[1]
        for b in todos_botoes
        if (abs(b[0] - x_btn_add) <= 180 or b[2] >= 112)
        and b[1] > y_piso
        and b[1] < 995
    ]
    y_lim = (min(proximos) - 32) if proximos else 1010
    if y_limite_extra is not None and int(y_limite_extra) > 220:
        y_lim = min(y_lim, int(y_limite_extra))
    return int(y_lim)


def _medir_scroll_por_template(
    img_antes: Optional[np.ndarray],
    img_depois: Optional[np.ndarray],
    offset_xy: Tuple[int, int],
    x_referencia: int,
    passos_scroll: int = -2,
) -> Optional[int]:
    """
    Mede com precisão de 1 pixel quantos pixels a página rolou verticalmente (delta_y = y_antes - y_depois)
    usando cv2.matchTemplate em toda a largura útil do formulário (X=680..1850), garantindo textura
    mesmo quando um dos lados do card tem poucas linhas.
    """
    if img_antes is None or img_depois is None:
        return None
    if img_antes.shape != img_depois.shape:
        return None

    ox, oy = offset_xy
    h_img, w_img = img_antes.shape[:2]

    # Usa toda a largura da área do formulário (X=680..1850) para ter sempre textura suficiente
    x1 = max(0, 680 - ox)
    x2 = min(w_img, 1850 - ox)
    if x2 - x1 < 120:
        x1 = max(0, min(w_img - 200, (w_img // 2) - 100))
        x2 = min(w_img, x1 + 200)

    if passos_scroll < 0:
        # Conteúdo vai subir: pega template na parte média-baixa (Y=500..700)
        y1 = max(210 - oy, min(h_img - 240, 500 - oy))
    else:
        # Conteúdo vai descer: pega template na parte média-alta (Y=220..420, abaixo da topbar fixa)
        y1 = max(210 - oy, min(h_img - 240, 220 - oy))
    y2 = min(h_img - 70, y1 + 200)
    if y2 - y1 < 80:
        return None

    gray_antes = cv2.cvtColor(img_antes[:, x1:x2], cv2.COLOR_BGR2GRAY)
    gray_depois = cv2.cvtColor(img_depois[:, x1:x2], cv2.COLOR_BGR2GRAY)

    tmpl = gray_antes[y1:y2, :]
    if float(np.std(tmpl)) < 2.5:
        return None

    y_busca_min = max(200 - oy, 0)
    y_busca_max = min(h_img, 995 - oy)
    if y_busca_max - y_busca_min <= (y2 - y1):
        return None

    area_busca = gray_depois[y_busca_min:y_busca_max, :]
    res = cv2.matchTemplate(area_busca, tmpl, cv2.TM_CCOEFF_NORMED)
    _, max_val, _, max_loc = cv2.minMaxLoc(res)
    if max_val < 0.72:
        return None

    y_novo = y_busca_min + max_loc[1]
    return int(y1 - y_novo)


def _parear_destaque_com_botao_adicionar(
    destaques: List[Tuple[int, int, int, int]],
    botoes_candidatos: List[Tuple[int, int, int, int]],
    incluir_botoes_largos: bool = False,
) -> Tuple[Optional[Tuple[int, int, int, int]], Optional[Tuple[int, int, int, int]]]:
    """
    Verifica qual palavra destacada pelo Ctrl+F tem um botão azul 'Adicionar' / 'Adicionar necessidade'
    na MESMA linha horizontal e dentro do MESMO card (dx entre 50 e 720px para cards de meia largura,
    ou até 1220px para blocos de largura total como 'Inventário de necessidades', dy <= 42px).
    """
    melhor_par: Tuple[Optional[Tuple[int, int, int, int]], Optional[Tuple[int, int, int, int]]] = (None, None)
    menor_score = 999999
    dx_max = 1220 if incluir_botoes_largos else 720
    dx_alvo = 950 if incluir_botoes_largos else 540

    for dest in destaques:
        hx, hy, _, _ = dest
        for btn in botoes_candidatos:
            bx, by, _, _ = btn
            dy = abs(by - hy)
            dx = bx - hx
            if dy <= 42 and 50 <= dx <= dx_max:
                score = (dy * 10) + abs(dx - dx_alvo)
                if score < menor_score:
                    menor_score = score
                    melhor_par = (btn, dest)

    return melhor_par


def _detectar_par_na_tela_atual(
    x_ref: int,
    x_btn_esperado: Optional[int],
    y_min_esperado: int,
    y_max_esperado: int,
    contexto_log: str = "",
    incluir_botoes_largos: bool = False,
) -> Tuple[
    Optional[Tuple[int, int, int, int]],
    Optional[Tuple[int, int, int, int]],
    int,
]:
    """
    Captura a tela atual e tenta parear o destaque do Ctrl+F com o botão 'Adicionar'
    da mesma coluna (nunca cruzando da coluna esquerda para a direita ou para outro monitor).
    """
    img, offset_xy = capturar_tela_sistema(x_ref)
    if img is None:
        return None, None, 0

    todos_botoes = [
        b
        for b in detectar_botoes_azuis_adicionar(
            img, offset_xy, incluir_botoes_largos=incluir_botoes_largos
        )
        if y_min_esperado <= b[1] <= y_max_esperado and 650 <= b[0] <= 1860
    ]
    if x_btn_esperado is not None:
        botoes_cand = [b for b in todos_botoes if abs(b[0] - x_btn_esperado) <= 240]
    else:
        botoes_cand = todos_botoes

    laranjas, amarelos = detectar_destaques_ctrl_f(img, offset_xy)

    def _destaque_na_coluna_esperada(d: Tuple[int, int, int, int]) -> bool:
        hx, hy, _, _ = d
        if not (y_min_esperado <= hy <= y_max_esperado and 600 <= hx <= 1850):
            return False
        if incluir_botoes_largos:
            return 600 <= hx <= 1850
        if x_btn_esperado is not None:
            if x_btn_esperado <= 1050:
                return 600 <= hx <= (x_btn_esperado - 35)
            if x_btn_esperado <= 1260:
                return 600 <= hx <= 1120
            if x_btn_esperado <= 1520:
                return 1020 <= hx <= (x_btn_esperado - 35)
            return 1150 <= hx <= (x_btn_esperado - 45)
        return True

    laranjas_validos = [d for d in laranjas if _destaque_na_coluna_esperada(d)]
    amarelos_validos = [
        d
        for d in amarelos
        if _destaque_na_coluna_esperada(d) and not (d[0] >= 1650 and d[1] <= 285)
    ]

    if contexto_log:
        _log_visao(
            f"{contexto_log}: laranjas={laranjas_validos}, "
            f"amarelos={amarelos_validos}, botoes={botoes_cand}"
        )

    btn_escolhido, dest_par = _parear_destaque_com_botao_adicionar(
        laranjas_validos, botoes_cand, incluir_botoes_largos=incluir_botoes_largos
    )
    if btn_escolhido is not None:
        return btn_escolhido, dest_par, len(laranjas_validos) + len(amarelos_validos)

    btn_escolhido, dest_par = _parear_destaque_com_botao_adicionar(
        amarelos_validos, botoes_cand, incluir_botoes_largos=incluir_botoes_largos
    )
    if btn_escolhido is not None:
        return btn_escolhido, dest_par, len(laranjas_validos) + len(amarelos_validos)

    destaque_escolhido: Optional[Tuple[int, int, int, int]] = None
    for h_cand in laranjas_validos + amarelos_validos:
        if h_cand[1] >= 145:
            destaque_escolhido = h_cand
            break

    return None, destaque_escolhido, len(laranjas_validos) + len(amarelos_validos)


def _buscar_campo_com_ctrl_f(
    termo_busca: str,
    x_btn_esperado: Optional[int] = None,
    y_min_esperado: int = 95,
    y_max_esperado: int = 1065,
    posicionar_no_topo: bool = True,
    espera_botao_adicionar: bool = True,
    incluir_botoes_largos: bool = False,
) -> Tuple[Optional[Tuple[int, int, int, int]], Optional[Tuple[int, int, int, int]]]:
    """
    Abre Ctrl+F, digita o nome exato do campo, localiza o destaque pareado com o botão 'Adicionar',
    posiciona o cabeçalho do card na área visível superior segura (Y ~ 215..430, abaixo da topbar fixa)
    e fecha o Ctrl+F.
    """
    params = obter_parametros_calibrados()
    xn, _ = params["ponto_neutro"]

    atalho_teclado("ctrl", "f")
    time.sleep(0.16)
    copiar_para_clipboard(termo_busca)
    time.sleep(0.06)
    atalho_teclado("ctrl", "a")
    time.sleep(0.04)
    atalho_teclado("ctrl", "v")
    time.sleep(0.32)

    x_ref = x_btn_esperado if x_btn_esperado is not None else max(950, xn)
    # Usa o corredor central do formulário (X=1230, Y=500) para que o scroll da roda sempre role a página
    x_scroll = 1230
    btn_escolhido: Optional[Tuple[int, int, int, int]] = None
    destaque_escolhido: Optional[Tuple[int, int, int, int]] = None

    for tentativa in range(5):
        btn_escolhido, destaque_escolhido, total_dest = _detectar_par_na_tela_atual(
            x_ref=x_ref,
            x_btn_esperado=x_btn_esperado,
            y_min_esperado=y_min_esperado,
            y_max_esperado=y_max_esperado,
            contexto_log=f"Ctrl+F '{termo_busca}' (tent={tentativa})",
            incluir_botoes_largos=incluir_botoes_largos,
        )
        if btn_escolhido is not None:
            break
        if not espera_botao_adicionar and total_dest == 1 and destaque_escolhido is not None:
            break
        # Se no 1º Ctrl+F o Chrome rolou para cima deixando o destaque escondido atrás da topbar fixa (Y < 198),
        # sobe a página +3 ticks para revelar o cabeçalho do card abaixo da topbar!
        if tentativa == 0 and total_dest == 0:
            rolar_pagina(3, x=x_scroll, y=500)
            time.sleep(0.32)
            continue
        if tentativa < 4:
            pressionar_tecla_unica("enter")
            time.sleep(0.25)

    pressionar_tecla_unica("esc")
    time.sleep(0.15)

    # Posiciona o cabeçalho do card em faixa segura (Y ~ 215..430), SEMPRE abaixo da topbar fixa (Y=198)
    if posicionar_no_topo and (btn_escolhido is not None or destaque_escolhido is not None):
        for ajuste in range(4):
            y_ref_atual = (
                btn_escolhido[1]
                if btn_escolhido is not None
                else (destaque_escolhido[1] if destaque_escolhido is not None else None)
            )
            if y_ref_atual is None:
                break

            # Se já está entre Y=215 e Y=430, está em posição ótima e segura (inclusive quando no fundo da página)
            if 215 <= y_ref_atual <= 430:
                break

            if y_ref_atual > 430:
                # Usa divisão inteira (// 100) mirando Y=250 para NUNCA jogar o botão para trás da topbar (< 208)
                passos = -max(1, min(7, int((y_ref_atual - 250) // 100)))
                print(
                    f"  📜 Posicionando '{termo_busca}' na tela "
                    f"(Y={y_ref_atual} -> scroll {passos})..."
                )
            elif y_ref_atual < 215:
                passos = max(1, min(5, int(((250 - y_ref_atual) + 80) // 100)))
                print(
                    f"  📜 Ajustando '{termo_busca}' na tela "
                    f"(Y={y_ref_atual} -> scroll +{passos})..."
                )
            else:
                break

            img_antes, off_antes = capturar_tela_sistema(x_ref)
            rolar_pagina(passos, x=x_scroll, y=500)
            time.sleep(0.38)
            img_depois, off_depois = capturar_tela_sistema(x_ref)

            delta_y = _medir_scroll_por_template(
                img_antes, img_depois, off_antes, x_ref, passos_scroll=passos
            )

            if btn_escolhido is not None and img_depois is not None:
                bx_ant, by_ant = int(btn_escolhido[0]), int(btn_escolhido[1])
                botoes_col = [
                    b
                    for b in detectar_botoes_azuis_adicionar(
                        img_depois, off_depois, incluir_botoes_largos=incluir_botoes_largos
                    )
                    if abs(b[0] - bx_ant) <= 95 and b[1] >= 205
                ]
                if delta_y is not None:
                    by_esp = by_ant - delta_y
                    if botoes_col:
                        botoes_col.sort(key=lambda b: abs(b[1] - by_esp))
                        if abs(botoes_col[0][1] - by_esp) <= 75:
                            btn_escolhido = botoes_col[0]
                            delta_y = by_ant - int(btn_escolhido[1])
                        else:
                            btn_escolhido = (bx_ant, by_esp, btn_escolhido[2], btn_escolhido[3])
                    else:
                        btn_escolhido = (bx_ant, by_esp, btn_escolhido[2], btn_escolhido[3])
                else:
                    by_est = by_ant - ((-passos) * 100)
                    if botoes_col:
                        if passos < 0:
                            cand_subiu = [b for b in botoes_col if b[1] <= (by_ant + 10)]
                            lista_busca = cand_subiu if cand_subiu else botoes_col
                        else:
                            cand_desceu = [b for b in botoes_col if b[1] >= (by_ant - 10)]
                            lista_busca = cand_desceu if cand_desceu else botoes_col
                        lista_busca.sort(key=lambda b: abs(b[1] - by_est))
                        btn_escolhido = lista_busca[0]
                        delta_y = by_ant - int(btn_escolhido[1])
                    else:
                        delta_y = (-passos) * 100
                        btn_escolhido = (bx_ant, by_est, btn_escolhido[2], btn_escolhido[3])

            if delta_y is None:
                delta_y = 0

            _log_visao(
                f"Pos-scroll '{termo_busca}' #{ajuste + 1}: passos={passos}, "
                f"delta_y={delta_y}px, btn_atual={btn_escolhido}"
            )

            if destaque_escolhido is not None:
                destaque_escolhido = (
                    destaque_escolhido[0],
                    int(destaque_escolhido[1]) - delta_y,
                    destaque_escolhido[2],
                    destaque_escolhido[3],
                )

            # Se a página atingiu o limite inferior/superior de rolagem (delta_y < 15px), encerra o ajuste!
            if abs(delta_y) < 15:
                break

    return btn_escolhido, destaque_escolhido


def _resolver_caixa1_do_card(
    x_btn_real: int,
    y_btn_real: int,
    destaque_titulo: Optional[Tuple[int, int, int, int]] = None,
    params: Optional[Dict[str, Any]] = None,
    coord_guia_caixa: Optional[Tuple[int, int]] = None,
) -> Tuple[int, int]:
    """
    Localiza a coordenada (X, Y) exata da 1ª CAIXA DE TEXTO VAZIA (<input>) abaixo do botão 'Adicionar'
    da coluna atual, usando a coordenada X do guia do mapeador (`coord_guia_caixa`) e a detecção
    visual de caixas vazias (`detectar_caixas_de_texto_da_coluna`).
    Garante que X NUNCA saia do monitor do sistema (650 <= X <= 1760).
    """
    if params is None:
        params = obter_parametros_calibrados()

    x_btn_real = max(650, min(1860, int(x_btn_real)))
    dx_btn_campo = int(params["dx_botao_para_campo"])
    dy_btn_campo = int(params["dy_botao_para_campo"])

    # Prioriza SEMPRE a coordenada horizontal X do guia apontado no mapeador para esta coluna
    if coord_guia_caixa is not None and 620 <= int(coord_guia_caixa[0]) <= 1780:
        x_caixa_real = int(coord_guia_caixa[0])
    elif destaque_titulo is not None and abs(destaque_titulo[1] - y_btn_real) <= 45:
        hx = int(destaque_titulo[0])
        x_meio_card = (hx + x_btn_real) // 2
        x_por_param = x_btn_real + dx_btn_campo
        x_caixa_real = max(hx + 35, min(x_btn_real - 110, (x_meio_card + x_por_param) // 2))
    else:
        x_caixa_real = x_btn_real + dx_btn_campo

    # Trava absoluta dentro do monitor do sistema (impossibilita clicar no 3º monitor X >= 1920)
    x_caixa_real = max(650, min(1760, int(x_caixa_real)))
    y_caixa1_real = y_btn_real + dy_btn_campo

    img, offset_xy = capturar_tela_sistema(x_btn_real)
    ys_vazias: List[int] = []
    ys_todas: List[int] = []
    if img is not None:
        todos_botoes = detectar_botoes_azuis_adicionar(
            img, offset_xy, incluir_botoes_largos=True
        )
        y_lim = _obter_limite_y_proximo_card(todos_botoes, x_btn_real, y_btn_real)
        ys_vazias = detectar_caixas_de_texto_da_coluna(
            img,
            offset_xy,
            x_guia_caixa=x_caixa_real,
            x_btn_add=x_btn_real,
            y_btn_add=y_btn_real,
            y_limite_inferior=y_lim,
            apenas_vazias=True,
        )
        ys_todas = detectar_caixas_de_texto_da_coluna(
            img,
            offset_xy,
            x_guia_caixa=x_caixa_real,
            x_btn_add=x_btn_real,
            y_btn_add=y_btn_real,
            y_limite_inferior=y_lim,
            apenas_vazias=False,
        )
        # 1. Prioriza SEMPRE uma caixa VAZIA do card atual (acima do limite do bloco vizinho y_lim)
        if ys_vazias:
            y_caixa1_real = ys_vazias[0]
        elif ys_todas:
            y_caixa1_real = ys_todas[0]

        if img is not None:
            y_caixa1_real = centralizar_y_na_caixa_branca(img, offset_xy, x_caixa_real, y_caixa1_real)

    _log_visao(
        f"_resolver_caixa1_do_card: btn=({x_btn_real}, {y_btn_real}), "
        f"guia={coord_guia_caixa}, ys_vazias={ys_vazias}, ys_todas={ys_todas} "
        f"-> caixa1=({x_caixa_real}, {y_caixa1_real})"
    )
    return (int(x_caixa_real), int(y_caixa1_real))


def localizar_card_com_ctrl_f_e_visao(
    nome_campo: str,
    coord_campo_salva: Optional[Tuple[int, int]] = None,
    coord_btn_salva: Optional[Tuple[int, int]] = None,
    eh_quadrante_inferior: bool = False,
    ja_rolou_para_baixo: bool = False,
    incluir_botoes_largos: bool = False,
    posicionar_no_topo: bool = True,
) -> Tuple[Tuple[int, int], Optional[Tuple[int, int]], bool]:
    """
    Localiza dinamicamente qualquer coluna/campo na página combinando:
    1. Guia do mapeador (`coord_campo_salva` = onde fica a caixa vazia inicial da coluna)
    2. Mapa HTML (título exato e estrutura do card)
    3. Ctrl+F no navegador + Visão OpenCV (botão azul 'Adicionar' / 'Adicionar necessidade' + caixa de texto vazia)
    4. Busca com rolagem progressiva caso os cards de cima tenham muitos itens e empurrem
       'Oportunidades', 'Ameaças' ou 'Inventário de necessidades' para baixo da tela.
    """
    params = obter_parametros_calibrados()
    mapa_html = obter_mapa_html_campos()

    chave_up = nome_campo.strip().upper()
    chave_norm = _norm_chave(nome_campo)
    if "INVENTARIO" in chave_norm or "NECESSIDADE" in chave_norm:
        incluir_botoes_largos = True

    info_html = mapa_html.get(chave_norm)
    if info_html and info_html.get("titulo_exato"):
        termo_busca = str(info_html["titulo_exato"])
    else:
        termo_busca = MAPA_TITULO_PAGINA.get(chave_up, nome_campo.strip())

    tem_btn_html = True if info_html is None else bool(info_html.get("tem_botao_adicionar", True))
    espera_botao = tem_btn_html or (coord_btn_salva is not None) or incluir_botoes_largos

    # Se temos o guia da caixa vazia inicial da coluna, deriva o X esperado do botão Adicionar
    eh_col_direita = chave_norm in ("FRAQUEZAS", "FRAQUEZA", "AMEACAS", "AMEACA") or (
        coord_campo_salva is not None and int(coord_campo_salva[0]) >= 1240
    )
    if coord_btn_salva is not None:
        x_btn_esperado: Optional[int] = max(650, min(1860, int(coord_btn_salva[0])))
    elif incluir_botoes_largos:
        x_btn_esperado = 1735
    elif coord_campo_salva is not None:
        x_btn_esperado = 1775 if int(coord_campo_salva[0]) >= 1240 else 1174
    else:
        x_btn_esperado = 1775 if eh_col_direita else 1174

    xn, yn = params["ponto_neutro"]
    xn_seguro = max(620, min(1800, int(xn)))
    yn_seguro = max(195, min(350, int(yn)))

    # Clique neutro leve para garantir foco na página antes do Ctrl+F
    mover_e_clicar(x=xn_seguro, y=yn_seguro, cliques=1, botao="left")
    time.sleep(0.10)

    btn_escolhido, destaque_escolhido = _buscar_campo_com_ctrl_f(
        termo_busca=termo_busca,
        x_btn_esperado=x_btn_esperado,
        posicionar_no_topo=posicionar_no_topo,
        espera_botao_adicionar=espera_botao,
        incluir_botoes_largos=incluir_botoes_largos,
    )

    # Se não encontrou o par (título + botão Adicionar) na tela atual, faz varredura com scroll progressivo!
    if (btn_escolhido is None and espera_botao) or (btn_escolhido is None and destaque_escolhido is None):
        sequencia_scroll = [-5, -5, -5, 8, 8] if eh_quadrante_inferior else [6, -5, -5]
        for idx_sc, passos_busca in enumerate(sequencia_scroll):
            _log_visao(
                f"Varredura de scroll para localizar '{termo_busca}' "
                f"(tentativa {idx_sc + 1}/{len(sequencia_scroll)}, scroll={passos_busca})"
            )
            rolar_pagina(passos_busca, x=1230, y=500)
            time.sleep(0.38)
            if passos_busca < 0:
                ja_rolou_para_baixo = True
            btn_escolhido, destaque_escolhido = _buscar_campo_com_ctrl_f(
                termo_busca=termo_busca,
                x_btn_esperado=x_btn_esperado,
                posicionar_no_topo=posicionar_no_topo,
                espera_botao_adicionar=espera_botao,
                incluir_botoes_largos=incluir_botoes_largos,
            )
            if btn_escolhido is not None or (not espera_botao and destaque_escolhido is not None):
                break

    # Se o Ctrl+F destacou o título na coluna certa mas o botão azul não foi pareado, busca na mesma linha
    if btn_escolhido is None and destaque_escolhido is not None:
        hx, hy, _, _ = destaque_escolhido
        if espera_botao:
            img_d, off_d = capturar_tela_sistema(hx)
            if img_d is not None:
                for b in detectar_botoes_azuis_adicionar(
                    img_d, off_d, incluir_botoes_largos=incluir_botoes_largos
                ):
                    if abs(b[1] - hy) <= 48 and b[0] > hx and abs(b[0] - x_btn_esperado) <= 240:
                        btn_escolhido = b
                        break
            if btn_escolhido is None:
                bx_est = max(650, min(1860, int(x_btn_esperado)))
                by_est = hy + int(params["dy_titulo_para_botao"])
                btn_escolhido = (bx_est, by_est, 84, 34)

    # Fallback visual por coluna caso o destaque do Ctrl+F não tenha aparecido
    if btn_escolhido is None:
        x_busca_col = x_btn_esperado if x_btn_esperado is not None else 1174
        img_col, off_col = capturar_tela_sistema(x_busca_col)
        if img_col is not None:
            todos_b = detectar_botoes_azuis_adicionar(
                img_col, off_col, incluir_botoes_largos=incluir_botoes_largos
            )
            botoes_col = [
                b
                for b in todos_b
                if (
                    b[0] > 1400
                    if (eh_col_direita or incluir_botoes_largos)
                    else b[0] <= 1400
                )
                and 205 <= b[1] <= 980
            ]
            if botoes_col:
                botoes_col.sort(key=lambda item: item[1])
                btn_escolhido = botoes_col[-1] if eh_quadrante_inferior else botoes_col[0]

    if btn_escolhido is not None:
        x_btn_real = max(650, min(1860, int(btn_escolhido[0])))
        y_btn_real = int(btn_escolhido[1])
        coord_caixa1 = _resolver_caixa1_do_card(
            x_btn_real=x_btn_real,
            y_btn_real=y_btn_real,
            destaque_titulo=destaque_escolhido,
            params=params,
            coord_guia_caixa=coord_campo_salva,
        )
        print(
            f"  🔍 [Guia + Visão] '{termo_busca}' localizado! "
            f"Botão Adicionar=({x_btn_real}, {y_btn_real}) | 1ª Caixa Vazia={coord_caixa1}"
        )
        return coord_caixa1, (x_btn_real, y_btn_real), ja_rolou_para_baixo

    if destaque_escolhido is not None:
        hx, hy, _, _ = destaque_escolhido
        x_campo_dinamico = (
            int(coord_campo_salva[0])
            if coord_campo_salva is not None
            else max(680, min(1760, hx + int(params["dx_titulo_para_campo"])))
        )
        x_campo_dinamico = max(650, min(1760, x_campo_dinamico))
        y_campo_dinamico = hy + int(params["dy_titulo_para_campo"])
        print(
            f"  🔍 [Guia + Visão] Campo '{termo_busca}' localizado em "
            f"({x_campo_dinamico}, {y_campo_dinamico})"
        )
        return (x_campo_dinamico, y_campo_dinamico), None, ja_rolou_para_baixo

    x_fallback = max(650, min(1760, int(coord_campo_salva[0]))) if coord_campo_salva else 910
    y_fallback = int(coord_campo_salva[1]) if coord_campo_salva else 488
    btn_fallback = (
        (max(650, min(1860, int(coord_btn_salva[0]))), int(coord_btn_salva[1]))
        if coord_btn_salva
        else None
    )
    return (x_fallback, y_fallback), btn_fallback, ja_rolou_para_baixo


def _calcular_deslocamento_real_scroll(ys_antes: List[int], ys_depois: List[int]) -> int:
    """
    Calcula o deslocamento vertical exato (em pixels) causado por um scroll da página
    comparando as posições dos botões '×' da mesma coluna antes e depois do scroll.
    """
    diffs: List[int] = []
    for ya in ys_antes:
        for yd in ys_depois:
            d = ya - yd
            if 30 <= d <= 360:
                diffs.append(d)

    if not diffs:
        return 200

    melhor_grupo: List[int] = []
    for d in diffs:
        grupo = [x for x in diffs if abs(x - d) <= 6]
        if len(grupo) > len(melhor_grupo):
            melhor_grupo = grupo

    escolhidos = melhor_grupo if melhor_grupo else diffs
    return int(round(sum(escolhidos) / float(len(escolhidos))))


def _esta_em_zona_proibida(
    x_alvo: int,
    y_alvo: int,
    zonas_proibidas: Optional[List[Tuple[int, int]]],
) -> bool:
    """Verifica se (x_alvo, y_alvo) colide com outra caixa mapeada/protegida da página."""
    if not zonas_proibidas:
        return False
    for zx, zy in zonas_proibidas:
        if abs(int(x_alvo) - int(zx)) <= 220 and abs(int(y_alvo) - int(zy)) <= 36:
            return True
    return False


def ajustar_coordenada_para_caixa_vazia(
    coord_alvo: Tuple[int, int],
    coord_btn_add: Optional[Tuple[int, int]] = None,
    zonas_proibidas: Optional[List[Tuple[int, int]]] = None,
    y_limite_bloco_vizinho: Optional[int] = None,
    permitir_avancar_abaixo: bool = True,
    nome_campo: Optional[str] = None,
) -> Tuple[Tuple[int, int], bool, bool]:
    """
    Garante antes de cada colagem que `(cx, cy)`:
    1. Permaneça estritamente dentro da `<div>` (`.wizard-table-card` / `data-repeater` / coluna Bootstrap) do campo.
    2. NÃO invada um bloco vizinho abaixo (`y_limite_bloco_vizinho` / próximo botão azul 'Adicionar...' / `zonas_proibidas`).
    3. Aponte para o centro vertical exato de uma CAIXA DE TEXTO VAZIA (`eh_caixa_de_texto_vazia`).
    4. Caso a caixa atual já esteja preenchida e exista outra caixa vazia mais abaixo na mesma coluna
       (acima do limite do bloco vizinho), avança automaticamente para a caixa vazia disponível.
    Retorna: `(coord_ajustada, esta_vazia, bloqueado_por_bloco_vizinho)`
    """
    regra_div = REGRAS_DIV_REPEATER_HTML.get(_norm_chave(nome_campo or "")) if nome_campo else None
    if regra_div is not None:
        x_min_div = int(regra_div.get("x_min_input", 650))
        x_max_div = int(regra_div.get("x_max_input", 1760))
        cx = max(x_min_div, min(x_max_div, int(coord_alvo[0])))
    else:
        cx = max(650, min(1760, int(coord_alvo[0])))
    cy = int(coord_alvo[1])

    if _esta_em_zona_proibida(cx, cy, zonas_proibidas):
        return ((cx, cy), False, True)
    if y_limite_bloco_vizinho is not None and int(y_limite_bloco_vizinho) > 220 and cy >= int(y_limite_bloco_vizinho):
        return ((cx, cy), False, True)

    img, offset_xy = capturar_tela_sistema(cx)
    if img is None:
        return ((cx, cy), True, False)

    todos_botoes = detectar_botoes_azuis_adicionar(img, offset_xy, incluir_botoes_largos=True)
    if regra_div is not None and regra_div.get("ultimo_bloco_da_pagina"):
        # Em <div class="wizard-table-card" data-repeater="needs"> (último bloco da página 3),
        # não existe outro repeater abaixo (apenas o rodapé da página).
        x_b = int(coord_btn_add[0]) if coord_btn_add is not None else cx
        y_b = max(65, int(coord_btn_add[1])) if coord_btn_add is not None else max(65, cy - 110)
        y_lim = int(y_limite_bloco_vizinho) if (y_limite_bloco_vizinho is not None and int(y_limite_bloco_vizinho) > 220) else 965
    elif coord_btn_add is not None:
        x_b = int(coord_btn_add[0])
        y_b_raw = int(coord_btn_add[1])
        y_b = max(65, y_b_raw)
        botoes_filtrados = todos_botoes
        if regra_div is not None and regra_div.get("bloqueado_por_botao_largo_abaixo"):
            # Oportunidades e Ameaças só têm abaixo deles o card 'Inventário de necessidades' (botão largo w >= 115px)
            botoes_filtrados = [b for b in todos_botoes if b[2] >= 115 and b[1] < 900]
        y_min_abaixo = (cy + 25) if y_b_raw < 215 else None
        y_lim = _obter_limite_y_proximo_card(
            botoes_filtrados,
            x_b,
            y_b,
            y_limite_extra=y_limite_bloco_vizinho,
            y_minimo_abaixo=y_min_abaixo,
        )
    else:
        x_b = cx
        y_b = max(65, cy - 110)
        y_lim = (
            int(y_limite_bloco_vizinho)
            if (y_limite_bloco_vizinho is not None and int(y_limite_bloco_vizinho) > 220)
            else 1010
        )

    if cy >= y_lim:
        return ((cx, cy), False, True)

    cy_cent = centralizar_y_na_caixa_branca(img, offset_xy, cx, cy)
    if cy_cent >= y_lim or _esta_em_zona_proibida(cx, cy_cent, zonas_proibidas):
        return ((cx, cy), False, True)

    vazia_atual = eh_caixa_de_texto_vazia(
        img, offset_xy, cx, cy_cent, x_b if coord_btn_add is not None else None
    )
    if vazia_atual:
        return ((cx, int(cy_cent)), True, False)

    if permitir_avancar_abaixo and coord_btn_add is not None:
        ys_vazias = detectar_caixas_de_texto_da_coluna(
            img,
            offset_xy,
            x_guia_caixa=cx,
            x_btn_add=x_b,
            y_btn_add=y_b,
            y_limite_inferior=y_lim,
            apenas_vazias=True,
        )
        vazias_abaixo = [
            yx
            for yx in ys_vazias
            if yx > (cy + 36)
            and yx < y_lim
            and not _esta_em_zona_proibida(cx, yx, zonas_proibidas)
        ]
        if vazias_abaixo:
            y_nova = centralizar_y_na_caixa_branca(img, offset_xy, cx, int(vazias_abaixo[0]))
            return ((cx, int(y_nova)), True, False)

    return ((cx, int(cy_cent)), False, False)


def _procurar_nova_caixa_surgida(
    img_depois: Optional[np.ndarray],
    offset_xy: Tuple[int, int],
    x_caixa: int,
    x_btn: int,
    y_btn: int,
    y_caixa_anterior: int,
    passo_y_padrao: int,
    zonas_proibidas: Optional[List[Tuple[int, int]]] = None,
) -> Tuple[int, List[int]]:
    """
    Procura visualmente na tela a nova CAIXA DE TEXTO VAZIA que acabou de surgir logo abaixo
    de `y_caixa_anterior` após clicar em 'Adicionar', isolando apenas os campos do card atual
    e evitando invadir caixas de outras seções/campos mapeados na página.
    """
    if img_depois is None:
        return y_caixa_anterior + passo_y_padrao, []

    todos_botoes = detectar_botoes_azuis_adicionar(
        img_depois, offset_xy, incluir_botoes_largos=True
    )
    y_lim = _obter_limite_y_proximo_card(todos_botoes, x_btn, max(70, y_btn))

    ys_todas_bruto = detectar_caixas_de_texto_da_coluna(
        img_depois,
        offset_xy,
        x_guia_caixa=x_caixa,
        x_btn_add=x_btn,
        y_btn_add=max(65, y_btn),
        y_limite_inferior=y_lim,
        apenas_vazias=False,
    )
    ys_todas = [
        yx
        for yx in ys_todas_bruto
        if yx < y_lim and not _esta_em_zona_proibida(x_caixa, yx, zonas_proibidas)
    ]

    # Calcula o passo vertical real desta coluna na escala/zoom atual do navegador (ex: 76px em 90%, 84px em 100%)
    diffs_col = [
        ys_todas[i] - ys_todas[i - 1]
        for i in range(1, len(ys_todas))
        if 52 <= (ys_todas[i] - ys_todas[i - 1]) <= 98
    ]
    passo_real = (
        int(round(float(np.median(diffs_col)))) if diffs_col else int(passo_y_padrao)
    )
    y_esperado = min(y_lim - 20, y_caixa_anterior + passo_real)

    ys_vazias = [
        yx
        for yx in ys_todas
        if eh_caixa_de_texto_vazia(img_depois, offset_xy, x_caixa, yx, x_btn)
    ]

    # 1. Prioriza uma caixa VAZIA estritamente 1 passo abaixo de y_caixa_anterior (45..112px)
    # Isso impede que o algoritmo pule +176px para dentro de outra caixa abaixo (ex: Inventário de necessidades)
    vazias_abaixo = [
        yx for yx in ys_vazias if (y_caixa_anterior + 45) <= yx <= (y_caixa_anterior + 112)
    ]
    if vazias_abaixo:
        vazias_abaixo.sort(key=lambda yx: abs(yx - y_esperado))
        y_esc = centralizar_y_na_caixa_branca(img_depois, offset_xy, x_caixa, int(vazias_abaixo[0]))
        return int(y_esc), ys_todas

    # 2. Se a caixa imediatamente abaixo já estava preenchida, pega a próxima caixa VAZIA disponível abaixo na mesma coluna
    vazias_mais_abaixo = [
        yx for yx in ys_vazias if yx > (y_caixa_anterior + 45) and yx < y_lim
    ]
    if vazias_mais_abaixo:
        y_esc = centralizar_y_na_caixa_branca(img_depois, offset_xy, x_caixa, int(vazias_mais_abaixo[0]))
        return int(y_esc), ys_todas

    # 3. Caso contrário, procura qualquer caixa com botão '×' da mesma coluna 1 passo abaixo (45..112px)
    candidatos_abaixo = [
        yx for yx in ys_todas if (y_caixa_anterior + 45) <= yx <= (y_caixa_anterior + 112)
    ]
    if candidatos_abaixo:
        candidatos_abaixo.sort(key=lambda yx: abs(yx - y_esperado))
        return int(candidatos_abaixo[0]), ys_todas

    return int(y_esperado), ys_todas


def clicar_adicionar_e_avancar_caixa(
    coord_caixa_atual: Tuple[int, int],
    coord_btn_atual: Tuple[int, int],
    passo_y_padrao: Optional[int] = None,
    zonas_proibidas: Optional[List[Tuple[int, int]]] = None,
    forcar_scroll_padrao: bool = False,
) -> Tuple[Tuple[int, int], Tuple[int, int], int]:
    """
    Executa o ciclo unitário dentro da coluna atual:
    1. Confirma a posição atual do botão 'Adicionar' da coluna e clica 1x nele
       (se o botão subiu para trás da barra fixa superior `.np-admin-topbar` em Y < 208,
       sobe a página, clica em 'Adicionar', aguarda e volta).
    2. Aguarda `ESPERA_APOS_ADICIONAR` (~1.25s) para a nova caixa surgir no DOM.
    3. Procura visualmente a nova CAIXA DE TEXTO VAZIA que surgiu logo abaixo do campo anterior,
       respeitando o passo real da coluna e as zonas protegidas de outras caixas da página.
    4. Padroniza um scroll para baixo na aba do sistema a cada 4 campos adicionados
       (`forcar_scroll_padrao=True`) ou quando o novo campo ficar perto da borda inferior (Y > 935),
       medindo o deslocamento exato em pixels e localizando a caixa vazia pós-scroll.
    Retorna: (nova_coord_caixa, novo_coord_btn, ticks_rolados_para_baixo)
    """
    params = obter_parametros_calibrados()
    if passo_y_padrao is None:
        passo_y_padrao = int(params["passo_y_nova_caixa"])

    x_caixa, y_caixa = int(coord_caixa_atual[0]), int(coord_caixa_atual[1])
    x_btn, y_btn = int(coord_btn_atual[0]), int(coord_btn_atual[1])
    x_scroll = max(900, min(1600, x_caixa))

    # 1. Verifica a posição exata do botão 'Adicionar' na tela antes de clicar
    img_antes, off_antes = capturar_tela_sistema(x_btn)
    if img_antes is not None:
        botoes_swot_antes = detectar_botoes_azuis_adicionar(
            img_antes, off_antes, incluir_botoes_largos=False
        )
        for b in botoes_swot_antes:
            if abs(b[0] - x_btn) <= 85 and abs(b[1] - y_btn) <= 45 and b[1] >= 205:
                x_btn, y_btn = int(b[0]), int(b[1])
                break

    # A barra superior fixa (.np-admin-topbar) ocupa até Y ~ 198; portanto o botão só é clicável em Y >= 208
    if y_btn >= 208:
        mover_e_clicar(x=x_btn, y=y_btn, cliques=1, botao="left")
        time.sleep(max(1.25, float(ESPERA_APOS_ADICIONAR)))
    else:
        # Coluna longa em que o botão 'Adicionar' ficou atrás da topbar fixa (Y < 208): sobe, clica, aguarda e volta
        passos_subir = max(2, min(6, int(round((300 - y_btn) / 100.0))))
        rolar_pagina(passos_subir, x=x_scroll, y=500)
        time.sleep(0.40)
        img_up, off_up = capturar_tela_sistema(x_btn)
        delta_up = _medir_scroll_por_template(
            img_antes, img_up, off_antes, x_btn, passos_scroll=passos_subir
        )
        if delta_up is None:
            delta_up = -(passos_subir * 100)
        x_clique_btn, y_clique_btn = x_btn, y_btn - delta_up
        if img_up is not None:
            botoes_up = [
                b
                for b in detectar_botoes_azuis_adicionar(
                    img_up, off_up, incluir_botoes_largos=False
                )
                if abs(b[0] - x_btn) <= 85 and 208 <= b[1] <= 580
            ]
            if botoes_up:
                botoes_up.sort(key=lambda b: abs(b[1] - y_clique_btn))
                x_clique_btn, y_clique_btn = int(botoes_up[0][0]), int(botoes_up[0][1])
        mover_e_clicar(x=x_clique_btn, y=y_clique_btn, cliques=1, botao="left")
        time.sleep(max(1.25, float(ESPERA_APOS_ADICIONAR)))
        rolar_pagina(-passos_subir, x=x_scroll, y=500)
        time.sleep(0.40)

    # 2. Captura a tela após aguardar a criação da nova caixa e procura a caixa vazia abaixo de y_caixa
    img_depois, off_depois = capturar_tela_sistema(x_btn)
    y_novo_detectado, ys_x_depois = _procurar_nova_caixa_surgida(
        img_depois=img_depois,
        offset_xy=off_depois,
        x_caixa=x_caixa,
        x_btn=x_btn,
        y_btn=y_btn,
        y_caixa_anterior=y_caixa,
        passo_y_padrao=passo_y_padrao,
        zonas_proibidas=zonas_proibidas,
    )

    deve_rolar_padrao = bool(forcar_scroll_padrao and y_novo_detectado >= 580)

    # 3. Se o novo campo está visível na tela (<= 935) e não é o 4º clique do ciclo padrão, retorna sem rolar
    if y_novo_detectado <= 935 and not deve_rolar_padrao:
        _log_visao(
            f"clicar_adicionar_e_avancar_caixa (sem scroll): nova_caixa=({x_caixa}, {y_novo_detectado}), "
            f"btn=({x_btn}, {y_btn}), ys_x={ys_x_depois}"
        )
        return (x_caixa, int(y_novo_detectado)), (x_btn, int(y_btn)), 0

    # 4. A cada 4 cliques em 'Adicionar' (`deve_rolar_padrao`) ou quando Y > 935, faz scroll down padronizado na aba do sistema
    if deve_rolar_padrao:
        ticks_rolar = 3 if y_novo_detectado >= 760 else 2
    else:
        ticks_rolar = 1 if y_novo_detectado <= 985 else 2
    rolar_pagina(-ticks_rolar, x=x_scroll, y=500)
    time.sleep(0.40)

    img_pos, off_pos = capturar_tela_sistema(x_btn)
    delta_real = _medir_scroll_por_template(
        img_depois, img_pos, off_depois, x_btn, passos_scroll=-ticks_rolar
    )

    if img_pos is not None:
        todos_b_pos = detectar_botoes_azuis_adicionar(
            img_pos, off_pos, incluir_botoes_largos=True
        )
        botoes_swot_pos = detectar_botoes_azuis_adicionar(
            img_pos, off_pos, incluir_botoes_largos=False
        )
        y_lim_pos = _obter_limite_y_proximo_card(
            todos_b_pos, x_btn, max(70, y_btn - (ticks_rolar * 100))
        )
        ys_x_pos = detectar_caixas_de_texto_da_coluna(
            img_pos,
            off_pos,
            x_guia_caixa=x_caixa,
            x_btn_add=x_btn,
            y_btn_add=60,
            y_limite_inferior=y_lim_pos,
            apenas_vazias=False,
        )

        if delta_real is None:
            if ys_x_depois and ys_x_pos:
                delta_real = _calcular_deslocamento_real_scroll(ys_x_depois, ys_x_pos)
            else:
                delta_real = ticks_rolar * 100

        y_caixa_anterior_pos = y_caixa - delta_real
        y_btn_pos = y_btn - delta_real

        for b in botoes_swot_pos:
            if abs(b[0] - x_btn) <= 85 and abs(b[1] - y_btn_pos) <= 45:
                y_btn_pos = int(b[1])
                break

        zonas_pos = (
            [(zx, zy - delta_real) for zx, zy in zonas_proibidas]
            if zonas_proibidas
            else None
        )
        y_caixa_pos, _ = _procurar_nova_caixa_surgida(
            img_depois=img_pos,
            offset_xy=off_pos,
            x_caixa=x_caixa,
            x_btn=x_btn,
            y_btn=y_btn_pos,
            y_caixa_anterior=y_caixa_anterior_pos,
            passo_y_padrao=passo_y_padrao,
            zonas_proibidas=zonas_pos,
        )
    else:
        if delta_real is None:
            delta_real = ticks_rolar * 100
        y_caixa_pos = y_novo_detectado - delta_real
        y_btn_pos = y_btn - delta_real

    _log_visao(
        f"clicar_adicionar_e_avancar_caixa (scroll -{ticks_rolar}, delta={delta_real}px): "
        f"nova_caixa=({x_caixa}, {y_caixa_pos}), btn=({x_btn}, {y_btn_pos})"
    )
    return (x_caixa, int(y_caixa_pos)), (x_btn, int(y_btn_pos)), ticks_rolar


def _localizar_botao_adicionar_necessidade_visivel(
    img_bgr: Optional[np.ndarray],
    offset_xy: Tuple[int, int],
    x_btn_ref: int,
    y_btn_esperado: Optional[int] = None,
    preferir_topo: bool = False,
) -> Optional[Tuple[int, int]]:
    """
    Localiza visualmente o botão azul do bloco ('Adicionar necessidade' ou 'Adicionar ação')
    na região visível abaixo do cabeçalho fixo (212 <= Y <= 860).
    - Para 'Inventário de necessidades' (`preferir_topo=False`), filtra w >= 112px e prioriza o mais abaixo.
    - Para 'Metas e ações' (`preferir_topo=True`), aceita w >= 90px e prioriza o botão do topo da página.
    """
    if img_bgr is None:
        return None
    todos_b = detectar_botoes_azuis_adicionar(
        img_bgr, offset_xy, incluir_botoes_largos=True
    )
    w_min = 90 if preferir_topo else 112
    candidatos_largos = [
        b
        for b in todos_b
        if b[2] >= w_min and abs(b[0] - x_btn_ref) <= 145 and 212 <= b[1] <= 860
    ]
    if candidatos_largos:
        if y_btn_esperado is not None and y_btn_esperado >= 212:
            perto_esp = [b for b in candidatos_largos if abs(b[1] - y_btn_esperado) <= 80]
            if perto_esp:
                perto_esp.sort(key=lambda b: abs(b[1] - y_btn_esperado))
                return int(perto_esp[0][0]), int(perto_esp[0][1])
            return None
        candidatos_largos.sort(key=lambda b: b[1], reverse=(not preferir_topo))
        return int(candidatos_largos[0][0]), int(candidatos_largos[0][1])

    # 2. Fallback caso o contorno azul do botão tenha sido segmentado menor, mas na posição exata esperada
    if y_btn_esperado is not None and y_btn_esperado >= 212:
        for b in todos_b:
            if abs(b[0] - x_btn_ref) <= 95 and abs(b[1] - y_btn_esperado) <= 40 and 212 <= b[1] <= 860:
                return int(b[0]), int(b[1])
    return None


def clicar_adicionar_e_avancar_grupo_codependente(
    coords_campos_atuais: List[Tuple[int, int]],
    coord_btn_atual: Tuple[int, int],
    passo_y_grupo: Optional[int] = None,
    zonas_proibidas: Optional[List[Tuple[int, int]]] = None,
    forcar_scroll_padrao: bool = False,
    indice_caixa: int = 1,
    coords_base_fundo: Optional[List[Tuple[int, int]]] = None,
    ancorado_no_topo: bool = False,
    texto_botao_adicionar: str = "Adicionar necessidade",
    y_limite_bloco_vizinho: Optional[int] = None,
    aprendizado_3_linhas: Optional[Dict[str, Any]] = None,
) -> Tuple[List[Tuple[int, int]], Tuple[int, int], int]:
    """
    Para grupos de campos codependentes na mesma linha/bloco:
    - Modo 1 (`ancorado_no_topo=False`, ex: 'Inventário de necessidades' no final da página 3):
      Sobe a página a partir da 2ª caixa para clicar em 'Adicionar necessidade' e rola até o fundo.
    - Modo 2 (`ancorado_no_topo=True`, ex: 'Metas e ações', 'Plano orçamentário e de investimentos'
      ou 'Plano de gestão de riscos' na página 4):
      Diferencia os 3 botões ('Adicionar ação', 'Adicionar linha', 'Adicionar risco') e o container
      `<div>` de cada bloco, clica no botão 'Adicionar...' do próprio bloco, desce de forma controlada
      parando ANTES da barreira do bloco vizinho (`_obter_limite_y_proximo_card`) e trava exclusivamente
      no sub-card VAZIO recém-criado (`eh_caixa_de_texto_vazia`).
    """
    if not coords_campos_atuais:
        return [], coord_btn_atual, 0

    base_coords = list(coords_base_fundo) if coords_base_fundo else list(coords_campos_atuais)
    x_btn = max(650, min(1860, int(coord_btn_atual[0])))
    y_btn = int(coord_btn_atual[1])
    x_scroll = 1230

    if ancorado_no_topo:
        norm_btn_add = _norm_chave(texto_botao_adicionar)
        eh_bloco_metas_topo = ("ACAO" in norm_btn_add or "METAS" in norm_btn_add)
        eh_bloco_orcamento = ("LINHA" in norm_btn_add or "ORCAMENT" in norm_btn_add)
        eh_bloco_riscos = ("RISCO" in norm_btn_add)
        eh_bloco_kpis = ("KPI" in norm_btn_add or "INDICADOR" in norm_btn_add)

        passo_padrao = 110 if eh_bloco_kpis else (188 if (eh_bloco_orcamento or eh_bloco_riscos) else 286)
        if eh_bloco_orcamento or eh_bloco_riscos:
            passo_card = 188 if (not passo_y_grupo or int(passo_y_grupo) >= 240) else int(passo_y_grupo)
        elif eh_bloco_metas_topo:
            passo_card = 286 if (not passo_y_grupo or int(passo_y_grupo) == 276) else int(passo_y_grupo)
        else:
            passo_card = int(passo_y_grupo) if (passo_y_grupo and int(passo_y_grupo) >= 85) else passo_padrao
        y_btn_calibrado = y_btn if y_btn >= 208 else (572 if eh_bloco_metas_topo else 300)

        img_antes, off_antes = capturar_tela_sistema(x_btn)

        # Caso especial Página 6 ('Indicadores de sucesso (KPIs)', data-repeater="kpis"):
        # O HTML inicial já vem com 3 linhas vazias pré-existentes (monitoring[kpis][0..2]).
        # Antes de clicar em 'Adicionar KPI', verifica se já existe uma linha vazia logo abaixo da linha atual!
        if eh_bloco_kpis and img_antes is not None:
            y_atual_kpi = int(coords_campos_atuais[-1][1])
            ys_x_kpi = detectar_botoes_remover_da_coluna(
                img_antes,
                off_antes,
                x_btn_add=x_btn,
                y_btn_add=max(65, y_btn),
                y_limite_inferior=965,
                passo_max_continuo=220,
            )
            proximas_kpi_vazias: List[int] = []
            for yx_k in ys_x_kpi:
                if (y_atual_kpi + 45) <= yx_k <= 945:
                    v_ind = eh_caixa_de_texto_vazia(
                        img_antes, off_antes, int(base_coords[0][0]), int(yx_k), None
                    )
                    v_form = eh_caixa_de_texto_vazia(
                        img_antes, off_antes, int(base_coords[-1][0]), int(yx_k), None
                    )
                    if v_ind and v_form:
                        proximas_kpi_vazias.append(int(yx_k))
            if proximas_kpi_vazias:
                proximas_kpi_vazias.sort()
                y_kpi_escolhido = proximas_kpi_vazias[0]
                ticks_kpi = 0
                if y_kpi_escolhido > 860:
                    ticks_kpi = 2
                    rolar_pagina(-ticks_kpi, x=x_scroll, y=500)
                    time.sleep(0.35)
                    img_k2, off_k2 = capturar_tela_sistema(x_btn)
                    delta_k = _medir_scroll_por_template(
                        img_antes, img_k2, off_antes, x_btn, passos_scroll=-ticks_kpi
                    )
                    if delta_k is None or delta_k <= 0 or delta_k > 260:
                        delta_k = ticks_kpi * 94
                    y_kpi_escolhido = int(y_kpi_escolhido - delta_k)
                    y_btn = int(y_btn - delta_k)
                    img_antes, off_antes = img_k2, off_k2
                novas_coords_kpi: List[Tuple[int, int]] = []
                for i in range(len(base_coords)):
                    cx_k = max(650, min(1760, int(base_coords[i][0])))
                    cy_k = int(y_kpi_escolhido)
                    if img_antes is not None:
                        cy_k = centralizar_y_na_caixa_branca(img_antes, off_antes, cx_k, cy_k)
                    novas_coords_kpi.append((cx_k, cy_k))
                _log_visao(
                    f"clicar_adicionar_e_avancar_grupo_codependente [kpis_linha_preexistente] "
                    f"(aba #{indice_caixa + 1}, sem clicar Adicionar KPI, scroll=-{ticks_kpi}): "
                    f"novas_coords={novas_coords_kpi}"
                )
                return novas_coords_kpi, (x_btn, y_btn), ticks_kpi

        btn_visivel = None
        if 220 <= y_btn <= 720:
            btn_visivel = _localizar_botao_adicionar_necessidade_visivel(
                img_antes,
                off_antes,
                x_btn,
                y_btn_esperado=y_btn,
                preferir_topo=True,
            )

        ticks_subidos = 0
        # Se o botão do bloco atual já está visível na tela (220 <= Y <= 720), clica direto sem subir!
        # Só sobe ou usa Ctrl+F se o botão saiu pelo topo da tela (btn_visivel is None ou Y < 225)
        if btn_visivel is None or btn_visivel[1] < 225:
            if eh_bloco_metas_topo:
                # Bloco 1 ('Metas e ações'): fica no topo absoluto da Página 4 (Y=572 com scroll=0)
                passos_iniciais_subir = max(4, (indice_caixa * 3) + 2)
                rolar_pagina(passos_iniciais_subir, x=x_scroll, y=500)
                ticks_subidos += passos_iniciais_subir
                time.sleep(0.38)
                img_up, off_up = capturar_tela_sistema(x_btn)
                btn_visivel = _localizar_botao_adicionar_necessidade_visivel(
                    img_up,
                    off_up,
                    x_btn,
                    y_btn_esperado=572,
                    preferir_topo=True,
                )

                for _ in range(18):
                    if btn_visivel is not None and 220 <= btn_visivel[1] <= 680:
                        break
                    img_prev, off_prev = capturar_tela_sistema(x_btn)
                    rolar_pagina(5, x=x_scroll, y=500)
                    ticks_subidos += 5
                    time.sleep(0.32)
                    img_up, off_up = capturar_tela_sistema(x_btn)
                    btn_visivel = _localizar_botao_adicionar_necessidade_visivel(
                        img_up,
                        off_up,
                        x_btn,
                        y_btn_esperado=572,
                        preferir_topo=True,
                    )
                    delta_up = _medir_scroll_por_template(
                        img_prev, img_up, off_prev, x_btn, passos_scroll=5
                    )
                    if delta_up is not None and abs(delta_up) < 10:
                        break
            else:
                # Blocos intermediários/finais ('Plano orçamentário e de investimentos' / 'Plano de gestão de riscos' / 'Indicadores de sucesso (KPIs)'):
                # NUNCA sobe até o topo da Página 4 (para não clicar em 'Adicionar ação' por engano)!
                # Posiciona o cabeçalho e o botão do próprio bloco em Y ~ 240..380 via Ctrl+F
                termo_cabecalho_bloco = (
                    "Plano orçamentário e de investimentos"
                    if eh_bloco_orcamento
                    else (
                        "Plano de gestão de riscos"
                        if eh_bloco_riscos
                        else (
                            "Indicadores de sucesso (KPIs)"
                            if eh_bloco_kpis
                            else texto_botao_adicionar
                        )
                    )
                )
                btn_cf, dest_cf = _buscar_campo_com_ctrl_f(
                    termo_busca=termo_cabecalho_bloco,
                    x_btn_esperado=x_btn,
                    posicionar_no_topo=True,
                    espera_botao_adicionar=True,
                    incluir_botoes_largos=True,
                )
                if btn_cf is not None:
                    btn_visivel = (int(btn_cf[0]), int(btn_cf[1]))
                elif dest_cf is not None:
                    btn_visivel = (x_btn, int(dest_cf[1]))

            if btn_visivel is None:
                atalho_teclado("ctrl", "f")
                time.sleep(0.18)
                copiar_para_clipboard(texto_botao_adicionar)
                time.sleep(0.06)
                atalho_teclado("ctrl", "a")
                time.sleep(0.04)
                atalho_teclado("ctrl", "v")
                time.sleep(0.35)
                pressionar_tecla_unica("esc")
                time.sleep(0.12)
                rolar_pagina(4, x=x_scroll, y=500)
                ticks_subidos += 4
                time.sleep(0.35)
                img_cf, off_cf = capturar_tela_sistema(x_btn)
                btn_visivel = _localizar_botao_adicionar_necessidade_visivel(
                    img_cf, off_cf, x_btn, y_btn_esperado=None, preferir_topo=True
                )
                if btn_visivel is None and img_cf is not None:
                    dest_cf = localizar_destaque_ctrl_f(img_cf, off_cf, ignorar_topo_y=185)
                    if dest_cf is not None:
                        btn_visivel = (int(dest_cf[0]), int(dest_cf[1]))

        if btn_visivel is not None:
            x_clique_btn, y_clique_btn = int(btn_visivel[0]), int(btn_visivel[1])
        else:
            x_clique_btn, y_clique_btn = x_btn, (572 if eh_bloco_metas_topo else max(240, min(600, y_btn_calibrado)))

        x_clique_btn = max(650, min(1860, int(x_clique_btn)))
        y_clique_btn = max(208, min(950, int(y_clique_btn)))
        mover_e_clicar(x=x_clique_btn, y=y_clique_btn, cliques=1, botao="left")
        time.sleep(max(1.25, float(ESPERA_APOS_ADICIONAR)))

        # Deslocamentos verticais relativos de cada campo até o último campo editável do sub-card (base_coords[-1])
        dy_rel_ult = [int(cy - base_coords[-1][1]) for _, cy in base_coords]
        dy_btn_para_ult_padrao = (
            115 if eh_bloco_kpis else (199 if (eh_bloco_orcamento or eh_bloco_riscos) else 289)
        )
        if indice_caixa == 1 and ticks_subidos == 0 and 208 <= y_btn_calibrado <= 700:
            dy_btn_para_ult = int(base_coords[-1][1] - y_btn_calibrado)
            if dy_btn_para_ult < 60 or dy_btn_para_ult > 360:
                dy_btn_para_ult = dy_btn_para_ult_padrao
        else:
            dy_btn_para_ult = dy_btn_para_ult_padrao

        # Diferença vertical entre o botão vermelho '×' do sub-card (yx) e o último campo editável (base_coords[-1][1]):
        # - Metas e ações: '×' está na Linha 2 (ao lado de PRAZO, base_coords[5]), enquanto RESPONSÁVEL está na Linha 3 (+89px abaixo do '×').
        # - Plano orçamentário: no HTML real (col-md-1 align-items-start pt-md-4), '×' está na Linha 1 (base_coords[0]), enquanto ANO 4 está na Linha 2 (+89px abaixo do '×').
        # - Plano de gestão de riscos: '×' está na Linha 1 (ao lado de NÍVEL, base_coords[0]), enquanto RESPONSÁVEL PELO RISCO está na Linha 2 (+89px abaixo do '×').
        # - Indicadores de sucesso (KPIs): '×' está na única linha horizontal (dy_x_para_ult = 0).
        if eh_bloco_kpis:
            dy_x_para_ult = 0
        elif eh_bloco_orcamento and len(base_coords) >= 7:
            dy_x_para_ult = int(base_coords[-1][1] - base_coords[0][1])
            if dy_x_para_ult < 45 or dy_x_para_ult > 135:
                dy_x_para_ult = 89
        elif eh_bloco_metas_topo and len(base_coords) >= 7:
            dy_x_para_ult = int(base_coords[-1][1] - base_coords[5][1])
            if dy_x_para_ult < 45 or dy_x_para_ult > 135:
                dy_x_para_ult = 89
        elif eh_bloco_riscos and len(base_coords) >= 5:
            dy_x_para_ult = int(base_coords[-1][1] - base_coords[0][1])
            if dy_x_para_ult < 45 or dy_x_para_ult > 135:
                dy_x_para_ult = 89
        else:
            dy_x_para_ult = 0

        # Se o usuário ensinou o comportamento de 3 linhas (--ensinar), usa o alvo aprendido para Linha 2 e Linha 3+
        y_alvo_resp = 785 if (eh_bloco_orcamento or eh_bloco_riscos or eh_bloco_kpis) else 845
        ticks_ensinado = None
        if isinstance(aprendizado_3_linhas, dict):
            if indice_caixa == 1 and isinstance(aprendizado_3_linhas.get("coords_linha2"), list) and aprendizado_3_linhas["coords_linha2"]:
                ult_c2 = aprendizado_3_linhas["coords_linha2"][-1]
                if isinstance(ult_c2, (list, tuple)) and len(ult_c2) >= 2:
                    y_alvo_resp = int(ult_c2[1])
                ticks_ensinado = aprendizado_3_linhas.get("ticks_scroll_linha2")
            elif indice_caixa >= 2 and isinstance(aprendizado_3_linhas.get("coords_linha3"), list) and aprendizado_3_linhas["coords_linha3"]:
                ult_c3 = aprendizado_3_linhas["coords_linha3"][-1]
                if isinstance(ult_c3, (list, tuple)) and len(ult_c3) >= 2:
                    y_alvo_resp = int(ult_c3[1])
                ticks_ensinado = aprendizado_3_linhas.get("ticks_scroll_linha3")

        ys_card_novo_antes = [
            int(y_clique_btn + dy_btn_para_ult + dy_rel_ult[i] + (indice_caixa * passo_card))
            for i in range(len(base_coords))
        ]
        y_resp_atual = ys_card_novo_antes[-1]

        delta_scroll_total = 0
        ticks_descidos_total = 0
        max_iter_scroll = max(4, indice_caixa + 4)

        for _ in range(max_iter_scroll):
            if y_resp_atual <= max(855, y_alvo_resp + 30):
                break
            px_faltantes = y_resp_atual - y_alvo_resp
            ticks_passo = max(1, min(4, int(round(px_faltantes / 95.0))))
            if ticks_descidos_total == 0 and ticks_ensinado is not None and int(ticks_ensinado) > 0:
                ticks_passo = max(1, min(6, int(ticks_ensinado)))
            elif indice_caixa == 1 and ticks_descidos_total == 0 and eh_bloco_metas_topo:
                ticks_passo = 3

            img_d0, off_d0 = capturar_tela_sistema(x_btn)
            rolar_pagina(-ticks_passo, x=x_scroll, y=500)
            ticks_descidos_total += ticks_passo
            time.sleep(0.38)
            img_d1, off_d1 = capturar_tela_sistema(x_btn)

            delta_step = _medir_scroll_por_template(
                img_d0, img_d1, off_d0, x_btn, passos_scroll=-ticks_passo
            )
            esperado_px = ticks_passo * 100
            if delta_step is not None and abs(delta_step) < 12:
                # Página encostou no fim absoluto da rolagem
                delta_step = 0
            elif delta_step is not None and 12 <= delta_step <= (esperado_px + 45) and (eh_bloco_riscos or eh_bloco_kpis):
                # No último bloco da página (ex: Riscos ou KPIs), aceita rolagem parcial ao encostar no rodapé
                delta_step = int(delta_step)
            elif delta_step is None or delta_step <= 0 or abs(delta_step - esperado_px) > 65:
                # Evita que sub-cards idênticos repetidos confundam o template matching em múltiplos de 286px
                delta_step = esperado_px

            delta_scroll_total += int(delta_step)
            y_resp_atual = ys_card_novo_antes[-1] - delta_scroll_total

            # Proteção contra invasão do bloco vizinho: se o botão azul do próximo bloco (ex: 'Adicionar linha'
            # ou 'Adicionar risco') já apareceu na parte inferior da tela (Y <= 890), para de descer imediatamente!
            if img_d1 is not None and not (eh_bloco_riscos or eh_bloco_kpis):
                todos_b_d1 = detectar_botoes_azuis_adicionar(img_d1, off_d1, incluir_botoes_largos=True)
                y_btn_topo_d1 = int(y_clique_btn - delta_scroll_total)
                y_lim_d1 = _obter_limite_y_proximo_card(
                    todos_b_d1,
                    x_btn,
                    max(70, y_btn_topo_d1),
                    y_limite_extra=y_limite_bloco_vizinho,
                    y_minimo_abaixo=max(380, int(y_btn_topo_d1) + 260),
                )
                if y_lim_d1 <= 890:
                    break

            if delta_step < 15:
                break

        # Refinamento visual pelo botão '×' + verificação de CAIXA VAZIA + BARREIRA DE BLOCO VIZINHO
        y_btn_pos_topo = int(y_clique_btn - delta_scroll_total)
        img_pos, off_pos = capturar_tela_sistema(x_btn)
        y_lim_vizinho = int(y_limite_bloco_vizinho) if (y_limite_bloco_vizinho and int(y_limite_bloco_vizinho) > 220) else 985

        if img_pos is not None:
            if eh_bloco_riscos or eh_bloco_kpis:
                # 'Plano de gestão de riscos' (Página 4) e 'Indicadores de sucesso (KPIs)' (Página 6)
                y_lim_vizinho = 965
            else:
                todos_b_pos = detectar_botoes_azuis_adicionar(img_pos, off_pos, incluir_botoes_largos=True)
                y_lim_vizinho = _obter_limite_y_proximo_card(
                    todos_b_pos,
                    x_btn,
                    max(70, y_btn_pos_topo),
                    y_limite_extra=y_limite_bloco_vizinho,
                    y_minimo_abaixo=max(380, int(y_btn_pos_topo) + 260),
                )
                # Se o bloco vizinho subiu demais na tela (y_lim_vizinho < 640), sobe para devolver o bloco atual ao centro
                if y_lim_vizinho < 640 and ticks_descidos_total >= 2:
                    ticks_recuo = min(ticks_descidos_total, max(2, int((680 - y_lim_vizinho) // 90)))
                    rolar_pagina(ticks_recuo, x=x_scroll, y=500)
                    ticks_descidos_total = max(0, ticks_descidos_total - ticks_recuo)
                    delta_scroll_total = max(0, delta_scroll_total - (ticks_recuo * 92))
                    y_btn_pos_topo = int(y_clique_btn - delta_scroll_total)
                    y_resp_atual = ys_card_novo_antes[-1] - delta_scroll_total
                    time.sleep(0.35)
                    img_pos, off_pos = capturar_tela_sistema(x_btn)
                    if img_pos is not None:
                        todos_b_pos = detectar_botoes_azuis_adicionar(img_pos, off_pos, incluir_botoes_largos=True)
                        y_lim_vizinho = _obter_limite_y_proximo_card(
                            todos_b_pos,
                            x_btn,
                            max(70, y_btn_pos_topo),
                            y_limite_extra=y_limite_bloco_vizinho,
                            y_minimo_abaixo=max(380, int(y_btn_pos_topo) + 260),
                        )

        # Se o cabeçalho/botão do próximo bloco abaixo (y_lim_vizinho) está visível na tela (480 <= y_lim_vizinho <= 945),
        # o novo sub-card recém-criado é SEMPRE o último sub-card do bloco atual, posicionado exatamente 62px acima de y_lim_vizinho!
        if not (eh_bloco_riscos or eh_bloco_kpis) and 480 <= y_lim_vizinho <= 945:
            y_resp_atual = int(y_lim_vizinho - 62)

        y_resp_escolhido = min(y_lim_vizinho - 22, int(y_resp_atual))
        if img_pos is not None:
            ys_x_pos = detectar_botoes_remover_da_coluna(
                img_pos,
                off_pos,
                x_btn_add=x_btn,
                y_btn_add=max(65, y_btn_pos_topo),
                y_limite_inferior=y_lim_vizinho,
                passo_max_continuo=380,
            )
            ys_x_pos = [
                yx
                for yx in ys_x_pos
                if yx < y_lim_vizinho
                and not _esta_em_zona_proibida(int(base_coords[-1][0]), int(yx + dy_x_para_ult), zonas_proibidas)
            ]

            # Mede a distância real entre sub-cards consecutivos pelos botões '×' visíveis (Lei da Distância Constante ΔY)
            if len(ys_x_pos) >= 2:
                ult_dy_x = int(ys_x_pos[-1] - ys_x_pos[-2])
                if abs(ult_dy_x - passo_card) <= 32:
                    passo_card = ult_dy_x

            # O novo sub-card recém-adicionado em um data-repeater é SEMPRE o último sub-card visível do bloco (ys_x_pos[-1])!
            ys_ult_todos = [
                int(yx + dy_x_para_ult)
                for yx in ys_x_pos
                if int(yx + dy_x_para_ult + dy_rel_ult[0]) >= 222
                and 222 <= int(yx + dy_x_para_ult) < (y_lim_vizinho - 10)
            ]
            if ys_ult_todos:
                ult_da_coluna = int(ys_ult_todos[-1])
                if abs(ult_da_coluna - y_resp_atual) <= max(105, int(passo_card * 0.45)):
                    y_resp_escolhido = ult_da_coluna

        novas_coords_topo: List[Tuple[int, int]] = []
        for i in range(len(base_coords)):
            cx_i = max(650, min(1760, int(base_coords[i][0])))
            cy_i = max(218, min(y_lim_vizinho - 18, int(y_resp_escolhido + dy_rel_ult[i])))
            if img_pos is not None:
                cy_i = centralizar_y_na_caixa_branca(img_pos, off_pos, cx_i, cy_i)
                cy_i = max(218, min(y_lim_vizinho - 18, int(cy_i)))
            novas_coords_topo.append((cx_i, cy_i))

        # Unifica o Y de todos os campos que pertencem à MESMA linha horizontal do sub-card
        # (ex: Linha 1 = [ID, OBJETIVO ESTRATÉGICO, META]; Linha 2 = [AÇÃO PRINCIPAL, INDICADOR, PRAZO]),
        # impedindo que o campo 'ID' (cujo label tem só 2 letras) fique desalinhado dos vizinhos da mesma linha!
        coords_alinhadas: List[Tuple[int, int]] = []
        for i, (cx_i, cy_i) in enumerate(novas_coords_topo):
            ys_mesma_linha = [
                novas_coords_topo[j][1]
                for j in range(len(base_coords))
                if abs(dy_rel_ult[j] - dy_rel_ult[i]) <= 12
            ]
            y_comum = int(round(float(np.median(ys_mesma_linha)))) if ys_mesma_linha else cy_i
            coords_alinhadas.append((cx_i, y_comum))
        novas_coords_topo = coords_alinhadas

        dy_snap = int(y_resp_escolhido - y_resp_atual)
        _log_visao(
            f"clicar_adicionar_e_avancar_grupo_codependente [ancorado_no_topo] (aba #{indice_caixa + 1}, "
            f"subiu=+{ticks_subidos}, btn_clicado=({x_clique_btn}, {y_clique_btn}), "
            f"desceu=-{ticks_descidos_total} [{delta_scroll_total}px], snap={dy_snap}px, y_lim_viz={y_lim_vizinho}): "
            f"novas_coords={novas_coords_topo}"
        )
        return novas_coords_topo, (x_btn, y_btn_pos_topo), ticks_descidos_total

    img_antes, off_antes = capturar_tela_sistema(x_btn)
    ys_x_antes: List[int] = []
    if img_antes is not None:
        ys_x_antes = detectar_botoes_remover_da_coluna(
            img_antes,
            off_antes,
            x_btn_add=x_btn,
            y_btn_add=65,
            y_limite_inferior=1010,
            passo_max_continuo=380,
        )

    btn_visivel = _localizar_botao_adicionar_necessidade_visivel(
        img_antes, off_antes, x_btn, y_btn_esperado=y_btn
    )

    # A cada 2 caixas (ou a partir da 2ª caixa / sempre que o botão subir para perto ou atrás do topo),
    # sobe a página ANTES de clicar em 'Adicionar necessidade' para localizar o botão com segurança!
    precisa_subir_para_achar_botao = (
        (indice_caixa >= 2)
        or (indice_caixa % 2 == 0)
        or (btn_visivel is None)
        or (btn_visivel[1] < 250)
    )

    ticks_subidos = 0
    x_clique_btn, y_clique_btn = x_btn, y_btn

    if not precisa_subir_para_achar_botao and btn_visivel is not None:
        x_clique_btn, y_clique_btn = int(btn_visivel[0]), int(btn_visivel[1])
    else:
        # Se o botão ainda está visível perto do topo (ex: 212..340), sobe um pouco (+3 ticks) para deixá-lo folgado
        if btn_visivel is not None and 212 <= btn_visivel[1] <= 360:
            rolar_pagina(3, x=x_scroll, y=500)
            ticks_subidos += 3
            time.sleep(0.35)
            img_up, off_up = capturar_tela_sistema(x_btn)
            btn_up = _localizar_botao_adicionar_necessidade_visivel(img_up, off_up, x_btn)
            if btn_up is not None:
                btn_visivel = btn_up
            else:
                btn_visivel = None

        # Se o botão já subiu acima da área visível (Y < 212), sobe a página em passos de +4 ticks (~380px)
        # até o botão 'Adicionar necessidade' entrar na tela (212 <= Y <= 860)
        if btn_visivel is None:
            if indice_caixa >= 6:
                salto_inicial = min(80, max(4, (indice_caixa - 3) * 2))
                rolar_pagina(salto_inicial, x=x_scroll, y=500)
                ticks_subidos += salto_inicial
                time.sleep(0.38)
                img_up, off_up = capturar_tela_sistema(x_btn)
                btn_visivel = _localizar_botao_adicionar_necessidade_visivel(img_up, off_up, x_btn)

            max_passos_busca = 28
            for _ in range(max_passos_busca):
                if btn_visivel is not None:
                    break
                rolar_pagina(4, x=x_scroll, y=500)
                ticks_subidos += 4
                time.sleep(0.32)
                img_up, off_up = capturar_tela_sistema(x_btn)
                btn_visivel = _localizar_botao_adicionar_necessidade_visivel(img_up, off_up, x_btn)

        # Fallback com Ctrl+F + subida de +4 ticks (para tirar o botão de trás do cabeçalho fixo!)
        if btn_visivel is None:
            atalho_teclado("ctrl", "f")
            time.sleep(0.20)
            copiar_para_clipboard(texto_botao_adicionar or "Adicionar necessidade")
            time.sleep(0.08)
            atalho_teclado("ctrl", "v")
            time.sleep(0.42)
            pressionar_tecla_unica("esc")
            time.sleep(0.12)
            # Sobe +4 ticks para que o botão 'Adicionar...' saia de trás da topbar fixa (Y < 200)
            rolar_pagina(4, x=x_scroll, y=500)
            ticks_subidos += 4
            time.sleep(0.38)
            img_cf, off_cf = capturar_tela_sistema(x_btn)
            btn_visivel = _localizar_botao_adicionar_necessidade_visivel(img_cf, off_cf, x_btn)
            if btn_visivel is None and img_cf is not None:
                dest_cf = localizar_destaque_ctrl_f(img_cf, off_cf, ignorar_topo_y=185)
                if dest_cf is not None:
                    btn_visivel = (int(dest_cf[0]), int(dest_cf[1]))

        if btn_visivel is not None:
            x_clique_btn, y_clique_btn = int(btn_visivel[0]), int(btn_visivel[1])
        else:
            x_clique_btn, y_clique_btn = x_btn, 360

    x_clique_btn = max(650, min(1860, int(x_clique_btn)))
    y_clique_btn = max(208, min(950, int(y_clique_btn)))
    mover_e_clicar(x=x_clique_btn, y=y_clique_btn, cliques=1, botao="left")
    time.sleep(max(1.25, float(ESPERA_APOS_ADICIONAR)))

    # Agora desce a página de volta até o final para posicionar os campos da nova caixa criada!
    ticks_descer = ticks_subidos + 6
    rolar_pagina(-ticks_descer, x=x_scroll, y=500)
    time.sleep(0.38)
    if ticks_subidos > 0 or indice_caixa >= 2:
        rolar_pagina(-8, x=x_scroll, y=500)
        time.sleep(0.32)
        ticks_descer += 8

    img_pos, off_pos = capturar_tela_sistema(x_btn)
    btn_pos = _localizar_botao_adicionar_necessidade_visivel(img_pos, off_pos, x_btn)
    y_btn_pos = int(btn_pos[1]) if btn_pos is not None else -100

    y_lim_fundo = int(y_limite_bloco_vizinho) if (y_limite_bloco_vizinho and int(y_limite_bloco_vizinho) > 220) else 1010
    dy_ajuste = 0
    if img_pos is not None and ys_x_antes:
        ys_x_pos = detectar_botoes_remover_da_coluna(
            img_pos,
            off_pos,
            x_btn_add=x_btn,
            y_btn_add=65,
            y_limite_inferior=y_lim_fundo,
            passo_max_continuo=380,
        )
        if ys_x_pos:
            diff_ult = int(ys_x_pos[-1] - ys_x_antes[-1])
            if abs(diff_ult) <= 35:
                dy_ajuste = diff_ult

    novas_coords_pos: List[Tuple[int, int]] = []
    for cx, cy in base_coords:
        cx_i = max(650, min(1760, int(cx)))
        cy_i = min(y_lim_fundo - 18, int(cy + dy_ajuste))
        if img_pos is not None:
            cy_i = centralizar_y_na_caixa_branca(img_pos, off_pos, cx_i, cy_i)
        novas_coords_pos.append((cx_i, cy_i))

    _log_visao(
        f"clicar_adicionar_e_avancar_grupo_codependente (caixa #{indice_caixa}, subiu=+{ticks_subidos}, "
        f"btn_clicado=({x_clique_btn}, {y_clique_btn}), desceu=-{ticks_descer}, ajuste={dy_ajuste}px): "
        f"novas_coords={novas_coords_pos}, novo_btn=({x_btn}, {y_btn_pos})"
    )
    return novas_coords_pos, (x_btn, y_btn_pos), 6


def obter_coordenada_caixa_item(
    indice_item: int,
    coord_caixa1: Tuple[int, int],
    coord_btn_add: Optional[Tuple[int, int]],
    passo_y_padrao: Optional[int] = None,
) -> Tuple[int, int]:
    """Compatibilidade: retorna a coordenada (X, Y) da caixa #(indice_item + 1)."""
    x_caixa1, y_caixa1 = int(coord_caixa1[0]), int(coord_caixa1[1])
    if indice_item <= 0 or coord_btn_add is None:
        return (x_caixa1, y_caixa1)
    if passo_y_padrao is None:
        passo_y_padrao = int(obter_parametros_calibrados()["passo_y_nova_caixa"])
    return (x_caixa1, min(995, y_caixa1 + (indice_item * passo_y_padrao)))


def selecionar_opcao_barra_suspensa(
    coord_caixa: Tuple[int, int],
    valor_opcao: str,
    opcoes_barra_suspensa: Optional[List[str]] = None,
    opcoes_coordenadas: Optional[Dict[str, List[int]]] = None,
) -> str:
    """
    Interage com uma barra suspensa de opções (<select class="form-select">) ou toggle:
    1. Clica na barra suspensa em `coord_caixa` para abrir/focar as opções.
    2. Se houver coordenadas de clique mapeadas por opção (`opcoes_coordenadas`), clica diretamente na opção.
    3. Caso contrário, seleciona a opção correspondente ('Alta', 'Média', 'Baixa') via tecla inicial + Enter.
    """
    x_sel, y_sel = int(coord_caixa[0]), int(coord_caixa[1])
    valor_limpo = str(valor_opcao).strip()
    chave_valor = _norm_chave(valor_limpo)

    # Caso a planilha tenha trazido uma nota numérica (ex: '4,33'), converte para Alta / Média / Baixa
    if not any(p in chave_valor for p in ("ALTA", "MEDIA", "BAIXA", "CRITICA")):
        try:
            num_val = float(valor_limpo.replace(".", "").replace(",", "."))
            if num_val >= 3.5:
                valor_limpo = "Alta"
                chave_valor = "ALTA"
            elif num_val >= 2.0:
                valor_limpo = "Média"
                chave_valor = "MEDIA"
            else:
                valor_limpo = "Baixa"
                chave_valor = "BAIXA"
        except Exception:
            pass

    # 1. Clica na barra suspensa / toggle para abrir o menu
    mover_e_clicar(x=x_sel, y=y_sel, cliques=1, botao="left")
    time.sleep(0.32)

    # 2. Se o usuário mapeou a coordenada de clique de cada opção aberta, clica nela diretamente
    if opcoes_coordenadas:
        for nome_op, coord_op in opcoes_coordenadas.items():
            op_norm = _norm_chave(nome_op)
            if (
                op_norm == chave_valor
                or (op_norm and op_norm in chave_valor)
                or (chave_valor and chave_valor in op_norm)
            ):
                if isinstance(coord_op, (list, tuple)) and len(coord_op) >= 2:
                    mover_e_clicar(x=int(coord_op[0]), y=int(coord_op[1]), cliques=1, botao="left")
                    time.sleep(0.25)
                    _log_visao(
                        f"selecionar_opcao_barra_suspensa (clique direto): '{valor_limpo}' -> ({coord_op[0]}, {coord_op[1]})"
                    )
                    return f"clique_opcao({nome_op})"

    # 3. Seleção em <select class="form-select"> (Alta / Média / Baixa):
    # Como 'Alta' começa com 'a', 'Média' com 'm' e 'Baixa' com 'b', pressionar a letra inicial
    # seleciona diretamente a opção correta mesmo quando há placeholder ('Selecione...') no topo!
    lista_opcoes = list(opcoes_barra_suspensa) if opcoes_barra_suspensa else ["Alta", "Média", "Baixa"]
    opcao_escolhida: Optional[str] = None
    for texto_op in lista_opcoes:
        op_norm = _norm_chave(texto_op)
        if op_norm == chave_valor or (op_norm and op_norm in chave_valor) or (chave_valor and chave_valor in op_norm):
            opcao_escolhida = texto_op
            break

    texto_alvo = opcao_escolhida if opcao_escolhida else valor_limpo
    chave_alvo = _norm_chave(texto_alvo)
    primeira_letra = next((ch.lower() for ch in chave_alvo if ch.isalnum()), "")
    if primeira_letra:
        pressionar_tecla_unica(primeira_letra)
        time.sleep(0.20)
    pressionar_tecla_unica("enter")
    time.sleep(0.25)
    _log_visao(
        f"selecionar_opcao_barra_suspensa (op='{texto_alvo}', tecla='{primeira_letra}'): caixa=({x_sel}, {y_sel})"
    )
    return f"tecla_{primeira_letra}({texto_alvo})"






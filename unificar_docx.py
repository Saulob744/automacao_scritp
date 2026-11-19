"""
unificar_docx.py
================
Unifica os 5 arquivos `.docx` gerados pelo Nex Privacy (Parte 1 a Parte 5, cada um salvo
abaixo do limite de 1.000 campos do servidor PHP) em um ÚNICO documento `.docx` oficial completo,
preservando 100% da formatação visual da Tecpar (capa, logo, cabeçalho, rodapé, estilos e bordas)
e recalculando automaticamente os totais consolidados do Plano Orçamentário!

Mapa de origem de cada seção:
- `parte1.docx` (Documento Base): Capa + Etapas 1 e 2 + 3.1 SWOT (117) + 3.2 Inventário (192)
- `parte2.docx`: 4.1 Metas e Ações (108 linhas)
- `parte3.docx` + `parte4.docx`: 4.2 Plano Orçamentário (linhas 1..95 + linhas 96..155 = 155 linhas + Totais recalculados)
- `parte5.docx`: 4.3 Gestão de Riscos (53) + 5.1 Capacitação (47) + 5.3 Ações Específicas (69) + 6.1 KPIs (108)
"""

from __future__ import annotations

import argparse
from copy import deepcopy
from pathlib import Path
import re
import sys
import unicodedata
from typing import Dict, List, Optional, Tuple

from docx import Document
from docx.oxml.ns import qn


PASTA_PADRAO_PARTES = Path(__file__).resolve().parent / "dados" / "partes_docx"
ARQUIVO_SAIDA_PADRAO = Path(__file__).resolve().parent / "PDTIC_Consolidado_Completo.docx"


def normalizar(txt: str) -> str:
    sem_acento = "".join(
        c for c in unicodedata.normalize("NFD", str(txt or "")) if unicodedata.category(c) != "Mn"
    )
    return re.sub(r"\s+", " ", sem_acento).strip().upper()


def extrair_numero_moeda(txt: str) -> float:
    s = str(txt or "").strip()
    if not s:
        return 0.0
    s = re.sub(r"[R$\s\u00a0]", "", s, flags=re.IGNORECASE)
    if re.match(r"^-?\d{1,3}(\.\d{3})*(,\d+)?$", s):
        s = s.replace(".", "").replace(",", ".")
    elif "," in s and "." not in s:
        s = s.replace(",", ".")
    else:
        s = re.sub(r"[^0-9.\-]", "", s)
    try:
        return float(s)
    except ValueError:
        return 0.0


def formatar_moeda_br(valor: float) -> str:
    s = f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


def texto_elemento_xml(el) -> str:
    textos = [t.text for t in el.iter(qn("w:t")) if t.text]
    return "".join(textos).strip()


def cabecalho_tabela_norm(tbl_el) -> str:
    trs = tbl_el.findall(qn("w:tr"))
    if not trs:
        return ""
    celulas = [normalizar(texto_elemento_xml(tc)) for tc in trs[0].findall(qn("w:tc"))]
    return " | ".join(celulas)


def classificar_tabela(tbl_el) -> Optional[str]:
    cab = cabecalho_tabela_norm(tbl_el)
    if not cab:
        return None
    if "PRIORIDADE" in cab and ("DEMANDANTE" in cab or "JUSTIFICATIVA" in cab):
        return "3.2_inventario"
    if "OBJETIVO" in cab and "META" in cab and ("ACAO" in cab or "INDICADOR" in cab):
        return "4.1_metas"
    if ("FONTE" in cab or "CUSTO" in cab) and ("ANO 1" in cab or "TOTAL" in cab or "2027" in cab):
        return "4.2_orcamento"
    if "RISCO" in cab and ("PROBABILIDADE" in cab or "IMPACTO" in cab or "MITIGACAO" in cab):
        return "4.3_riscos"
    if "INDICADOR" in cab and ("FORMULA" in cab or "DESCRICAO" in cab or "META" in cab) and "OBJETIVO" not in cab:
        return "6.1_kpis"
    return None


def detectar_codigo_secao_paragrafo(p_el) -> Optional[str]:
    """Detecta títulos numerados do PDTIC como '3.1', '4.1', '4.2', '4.3', '5.1', '5.2', '5.3', '6.'."""
    txt = normalizar(texto_elemento_xml(p_el))
    if not txt or len(txt) > 140:
        return None
    m_sub = re.match(r"^(\d+\.\d+)\b", txt)
    if m_sub:
        cod = m_sub.group(1)
        if cod in {"3.1", "3.2", "4.1", "4.2", "4.3", "5.1", "5.2", "5.3", "6.1", "6.2", "6.3"}:
            return cod
    m_cap = re.match(r"^([1-6])[\.\s\-—:]+\s*(IDENTIFICACAO|GOVERNANCA|DIAGNOSTICO|METAS|CAPACITACAO|MONITORAMENTO|INDICADORES)", txt)
    if m_cap:
        return f"{m_cap.group(1)}.0"
    return None


def mapear_intervalos_secoes(body_el) -> Dict[str, Tuple[int, int]]:
    """Mapeia `(idx_inicio, idx_fim_exclusivo)` de cada seção numerada nos filhos diretos de `<w:body>`."""
    filhos = list(body_el)
    marcadores: List[Tuple[int, str]] = []
    idx_fim_util = len(filhos)
    for i, el in enumerate(filhos):
        if el.tag == qn("w:sectPr"):
            idx_fim_util = i
            break
        if el.tag == qn("w:p"):
            cod = detectar_codigo_secao_paragrafo(el)
            if cod and not any(c_ant == cod for _, c_ant in marcadores):
                marcadores.append((i, cod))

    intervalos: Dict[str, Tuple[int, int]] = {}
    for pos_m, (idx_ini, cod) in enumerate(marcadores):
        # A subseção termina quando começa qualquer próxima seção/subseção
        idx_fim = idx_fim_util
        for idx_prox, cod_prox in marcadores[pos_m + 1:]:
            if cod.endswith(".0"):
                # Capítulo inteiro só termina no próximo capítulo X.0
                if cod_prox.endswith(".0") and cod_prox[0] != cod[0]:
                    idx_fim = idx_prox
                    break
            else:
                idx_fim = idx_prox
                break
        intervalos[cod] = (idx_ini, idx_fim)
    return intervalos


def eh_linha_total_orcamento(tr_el) -> bool:
    txt_tr = normalizar(texto_elemento_xml(tr_el))
    celulas = [normalizar(texto_elemento_xml(tc)) for tc in tr_el.findall(qn("w:tc"))]
    if not celulas:
        return False
    primeiras = " ".join(celulas[:3])
    return bool("TOTAL GERAL" in txt_tr or "TOTAIS" in primeiras or celulas[0] in ("TOTAL", "TOTAL GERAL"))


def atualizar_texto_celula_preservando_estilo(tc_el, novo_texto: str) -> None:
    runs_t = list(tc_el.iter(qn("w:t")))
    if runs_t:
        runs_t[0].text = novo_texto
        for extra in runs_t[1:]:
            extra.text = ""
        return
    p_el = tc_el.find(qn("w:p"))
    if p_el is not None:
        from docx.oxml import OxmlElement

        r_el = OxmlElement("w:r")
        t_el = OxmlElement("w:t")
        t_el.text = novo_texto
        r_el.append(t_el)
        p_el.append(r_el)


def mesclar_tabelas_orcamento(tbl_p3, tbl_p4) -> Tuple[object, Dict[str, float], int]:
    """Combina as linhas de dados da Parte 3 (1..95) + Parte 4 (96..155) e recalcula os totais."""
    tbl_final = deepcopy(tbl_p3)
    trs_final = tbl_final.findall(qn("w:tr"))
    trs_p4 = tbl_p4.findall(qn("w:tr"))

    linha_rodape_total = None
    if len(trs_final) > 1 and eh_linha_total_orcamento(trs_final[-1]):
        linha_rodape_total = deepcopy(trs_final[-1])
        tbl_final.remove(trs_final[-1])

    # Copia as linhas de dados da Parte 4 (pula cabeçalho idx=0 e eventual linha de total final)
    adicionadas_p4 = 0
    for tr in trs_p4[1:]:
        if eh_linha_total_orcamento(tr):
            continue
        tbl_final.append(deepcopy(tr))
        adicionadas_p4 += 1

    # Recalcula somas das colunas monetárias (Ano 1, Ano 2, Ano 3, Ano 4, Total)
    trs_dados = tbl_final.findall(qn("w:tr"))[1:]
    cab_celulas = [normalizar(texto_elemento_xml(tc)) for tc in tbl_final.findall(qn("w:tr"))[0].findall(qn("w:tc"))]

    indices_col_ano: Dict[str, int] = {}
    for idx_c, nome_c in enumerate(cab_celulas):
        if "ANO 1" in nome_c or "2027" in nome_c:
            indices_col_ano["ano1"] = idx_c
        elif "ANO 2" in nome_c or "2028" in nome_c:
            indices_col_ano["ano2"] = idx_c
        elif "ANO 3" in nome_c or "2029" in nome_c:
            indices_col_ano["ano3"] = idx_c
        elif "ANO 4" in nome_c or "2030" in nome_c:
            indices_col_ano["ano4"] = idx_c
        elif "TOTAL" in nome_c:
            indices_col_ano["total"] = idx_c

    somas = {"ano1": 0.0, "ano2": 0.0, "ano3": 0.0, "ano4": 0.0, "total": 0.0}
    for tr in trs_dados:
        tcs = tr.findall(qn("w:tc"))
        for chave_ano, idx_c in indices_col_ano.items():
            if idx_c < len(tcs):
                somas[chave_ano] += extrair_numero_moeda(texto_elemento_xml(tcs[idx_c]))

    if somas["total"] == 0.0:
        somas["total"] = somas["ano1"] + somas["ano2"] + somas["ano3"] + somas["ano4"]

    if linha_rodape_total is not None:
        tcs_tot = linha_rodape_total.findall(qn("w:tc"))
        for chave_ano, idx_c in indices_col_ano.items():
            if idx_c < len(tcs_tot):
                atualizar_texto_celula_preservando_estilo(tcs_tot[idx_c], formatar_moeda_br(somas[chave_ano]))
        tbl_final.append(linha_rodape_total)

    return tbl_final, somas, len(trs_dados)


def atualizar_paragrafos_totais_orcamento(elementos: List[object], somas: Dict[str, float]) -> int:
    """Se houver parágrafo resumo de 'Totais consolidados Ano 1: R$ ...', atualiza com a soma real das 155 linhas."""
    atualizados = 0
    for el in elementos:
        for t_el in el.iter(qn("w:t")):
            if not t_el.text:
                continue
            txt = t_el.text
            if "Ano 1:" in txt and "Total" in txt:
                novo = txt
                novo = re.sub(r"Ano\s*1:\s*R\$\s*[\d\.,]+", f"Ano 1: {formatar_moeda_br(somas['ano1'])}", novo)
                novo = re.sub(r"Ano\s*2:\s*R\$\s*[\d\.,]+", f"Ano 2: {formatar_moeda_br(somas['ano2'])}", novo)
                novo = re.sub(r"Ano\s*3:\s*R\$\s*[\d\.,]+", f"Ano 3: {formatar_moeda_br(somas['ano3'])}", novo)
                novo = re.sub(r"Ano\s*4:\s*R\$\s*[\d\.,]+", f"Ano 4: {formatar_moeda_br(somas['ano4'])}", novo)
                novo = re.sub(
                    r"Total(\s+geral)?:\s*R\$\s*[\d\.,]+",
                    lambda m: f"Total{m.group(1) or ''}: {formatar_moeda_br(somas['total'])}",
                    novo,
                    flags=re.IGNORECASE,
                )
                t_el.text = novo
                atualizados += 1
    return atualizados


def substituir_fatia_body(body_destino, ini: int, fim: int, novos_elementos: List[object]) -> None:
    antigos = list(body_destino)[ini:fim]
    if not antigos:
        return
    ancora = antigos[0]
    for novo in novos_elementos:
        ancora.addprevious(deepcopy(novo))
    for ant in antigos:
        body_destino.remove(ant)


def inspecionar_docx(caminho: Path) -> None:
    doc = Document(str(caminho))
    body = doc.element.body
    print("\n" + "=" * 90)
    print(f"🔬 INSPEÇÃO ESTRUTURAL DO ARQUIVO DOCX: {caminho.name}")
    print("=" * 90)
    intervalos = mapear_intervalos_secoes(body)
    print(f"📌 Seções numeradas detectadas ({len(intervalos)}):")
    for cod, (ini, fim) in intervalos.items():
        titulo_txt = texto_elemento_xml(list(body)[ini])[:75]
        print(f"   • [{cod}] elementos {ini}..{fim} ({fim - ini} bloco(s)) -> \"{titulo_txt}\"")

    print("\n📊 Tabelas detectadas:")
    for idx, tbl in enumerate(body.findall(qn("w:tbl"))):
        trs = tbl.findall(qn("w:tr"))
        tipo = classificar_tabela(tbl) or "outra"
        cab = cabecalho_tabela_norm(tbl)[:85]
        print(f"   • Tabela #{idx + 1} [{tipo}]: {max(0, len(trs) - 1)} linha(s) de dados | Cabeçalho: {cab}")
    print("=" * 90 + "\n")


def localizar_arquivo_parte(num_parte: int, caminho_explicito: Optional[str], pasta_partes: Path) -> Optional[Path]:
    if caminho_explicito:
        p = Path(caminho_explicito)
        return p if p.is_file() else None

    candidatos = [
        pasta_partes / f"parte{num_parte}.docx",
        pasta_partes / f"parte_{num_parte}.docx",
        pasta_partes / f"Parte {num_parte}.docx",
        Path.cwd() / f"parte{num_parte}.docx",
        Path.home() / "Downloads" / f"parte{num_parte}.docx",
    ]
    for c in candidatos:
        if c.is_file():
            return c
    return None


def unificar_5_partes(
    caminhos: Dict[int, Path],
    saida: Path,
) -> Path:
    print("\n" + "=" * 90)
    print("🧩 UNIFICADOR OFICIAL PDTIC (.DOCX) — JUNTANDO PARTES 1 A 5")
    print("=" * 90)
    for k in range(1, 6):
        print(f"   📄 Parte {k}: {caminhos[k]}")

    docs = {k: Document(str(caminhos[k])) for k in range(1, 6)}
    doc_final = docs[1]
    body_final = doc_final.element.body

    # 1. Substituir seções completas por fatiamento de <w:body> quando os títulos numerados existirem,
    #    com fallback direto por tabela classificada!
    secoes_origem = [
        ("4.1", 2),
        ("4.3", 5),
        ("5.1", 5),
        ("5.3", 5),
        ("6.1", 5),
    ]

    for cod_sec, num_doc_orig in secoes_origem:
        mapa_dest = mapear_intervalos_secoes(body_final)
        mapa_orig = mapear_intervalos_secoes(docs[num_doc_orig].element.body)
        if cod_sec in mapa_dest and cod_sec in mapa_orig:
            d_ini, d_fim = mapa_dest[cod_sec]
            o_ini, o_fim = mapa_orig[cod_sec]
            fatia_orig = list(docs[num_doc_orig].element.body)[o_ini:o_fim]
            substituir_fatia_body(body_final, d_ini, d_fim, fatia_orig)
            print(f"   ✅ Seção {cod_sec} copiada integralmente da Parte {num_doc_orig} ({len(fatia_orig)} elementos XML).")

    # Fallback/garantia adicional para tabelas 4.1, 4.3 e 6.1 caso alguma seção não use prefixo numérico padrão
    for tipo_tbl, num_doc_orig in [("4.1_metas", 2), ("4.3_riscos", 5), ("6.1_kpis", 5)]:
        tbl_dest = next((t for t in body_final.findall(qn("w:tbl")) if classificar_tabela(t) == tipo_tbl), None)
        tbl_orig = next((t for t in docs[num_doc_orig].element.body.findall(qn("w:tbl")) if classificar_tabela(t) == tipo_tbl), None)
        if tbl_dest is not None and tbl_orig is not None:
            qtd_dest = len(tbl_dest.findall(qn("w:tr"))) - 1
            qtd_orig = len(tbl_orig.findall(qn("w:tr"))) - 1
            if qtd_orig > qtd_dest:
                tbl_dest.addprevious(deepcopy(tbl_orig))
                tbl_dest.getparent().remove(tbl_dest)
                print(f"   ✅ Tabela [{tipo_tbl}] atualizada da Parte {num_doc_orig} ({qtd_orig} linhas).")

    # 2. Mesclar Seção 4.2 (Plano Orçamentário = Parte 3 [linhas 1..95] + Parte 4 [linhas 96..155])
    mapa_dest = mapear_intervalos_secoes(body_final)
    mapa_p3 = mapear_intervalos_secoes(docs[3].element.body)
    if "4.2" in mapa_dest and "4.2" in mapa_p3:
        d_ini, d_fim = mapa_dest["4.2"]
        o_ini, o_fim = mapa_p3["4.2"]
        substituir_fatia_body(body_final, d_ini, d_fim, list(docs[3].element.body)[o_ini:o_fim])

    tbl_orc_dest = next((t for t in body_final.findall(qn("w:tbl")) if classificar_tabela(t) == "4.2_orcamento"), None)
    tbl_orc_p3 = next((t for t in docs[3].element.body.findall(qn("w:tbl")) if classificar_tabela(t) == "4.2_orcamento"), None)
    tbl_orc_p4 = next((t for t in docs[4].element.body.findall(qn("w:tbl")) if classificar_tabela(t) == "4.2_orcamento"), None)

    if tbl_orc_dest is not None and tbl_orc_p3 is not None and tbl_orc_p4 is not None:
        tbl_unificada, somas_orc, tot_linhas_orc = mesclar_tabelas_orcamento(tbl_orc_p3, tbl_orc_p4)
        tbl_orc_dest.addprevious(tbl_unificada)
        tbl_orc_dest.getparent().remove(tbl_orc_dest)
        atualizar_paragrafos_totais_orcamento(list(body_final), somas_orc)
        print(
            f"   ✅ Tabela [4.2 Plano Orçamentário] unificada (Parte 3 + Parte 4) -> {tot_linhas_orc} linhas | "
            f"Ano 1={formatar_moeda_br(somas_orc['ano1'])} | Ano 2={formatar_moeda_br(somas_orc['ano2'])} | "
            f"Total Geral={formatar_moeda_br(somas_orc['total'])}!"
        )
    else:
        print("   ⚠️ Aviso: Tabela 4.2 não localizada automaticamente em uma das partes; verifique com `--inspecionar`.")

    saida.parent.mkdir(parents=True, exist_ok=True)
    doc_final.save(str(saida))
    print("-" * 90)
    print(f"🎉 ARQUIVO ÚNICO GERADO COM SUCESSO: {saida}")
    print("=" * 90 + "\n")
    inspecionar_docx(saida)
    return saida


def main() -> None:
    parser = argparse.ArgumentParser(description="Unifica parte1.docx .. parte5.docx em um único PDTIC oficial.")
    parser.add_argument("--pasta", type=str, default=str(PASTA_PADRAO_PARTES), help="Pasta contendo parte1.docx a parte5.docx.")
    parser.add_argument("--p1", type=str, default=None, help="Caminho explícito da Parte 1 (.docx).")
    parser.add_argument("--p2", type=str, default=None, help="Caminho explícito da Parte 2 (.docx).")
    parser.add_argument("--p3", type=str, default=None, help="Caminho explícito da Parte 3 (.docx).")
    parser.add_argument("--p4", type=str, default=None, help="Caminho explícito da Parte 4 (.docx).")
    parser.add_argument("--p5", type=str, default=None, help="Caminho explícito da Parte 5 (.docx).")
    parser.add_argument("--saida", "-o", type=str, default=str(ARQUIVO_SAIDA_PADRAO), help="Caminho do .docx unificado de saída.")
    parser.add_argument("--inspecionar", type=str, default=None, help="Apenas inspeciona as seções e tabelas de um arquivo .docx.")
    args = parser.parse_args()

    if args.inspecionar:
        arq = Path(args.inspecionar)
        if not arq.is_file():
            print(f"❌ Arquivo não encontrado: {arq}")
            sys.exit(1)
        inspecionar_docx(arq)
        return

    pasta = Path(args.pasta)
    explicitos = {1: args.p1, 2: args.p2, 3: args.p3, 4: args.p4, 5: args.p5}
    encontrados: Dict[int, Path] = {}
    faltando: List[int] = []
    for num in range(1, 6):
        arq_p = localizar_arquivo_parte(num, explicitos[num], pasta)
        if arq_p is None:
            faltando.append(num)
        else:
            encontrados[num] = arq_p

    if faltando:
        print("\n❌ Ainda faltam os seguintes arquivos .docx para concluir a unificação:")
        for num in faltando:
            print(f"   • parte{num}.docx (coloque em `{pasta}` ou em `Downloads`)")
        sys.exit(1)

    unificar_5_partes(encontrados, Path(args.saida))


if __name__ == "__main__":
    main()

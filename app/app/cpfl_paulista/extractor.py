# -*- coding: utf-8 -*-
"""
cpfl_paulista/extractor.py
Parser de faturas CPFL Paulista — Tarifa Verde A4 (média tensão, com demanda)
"""
import re
from pathlib import Path

try:
    import pdfplumber
except ImportError:
    pdfplumber = None


def _num(s):
    if not s:
        return None
    s = str(s).strip().replace('.', '').replace(',', '.')
    try:
        return float(s)
    except:
        return None


def _data(s):
    return s.strip() if s else None


def parse_fatura(pdf_path):
    if pdfplumber is None:
        raise ImportError("pdfplumber não instalado")

    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join(p.extract_text() or "" for p in pdf.pages)

    r = {"arquivo": Path(pdf_path).name, "distribuidora": "CPFL Paulista"}

    # ── Ref mês/ano + vencimento + total ─────────────────────────────────────
    m = re.search(
        r'(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)/(\d{4})\s+'
        r'(\d{2}/\d{2}/\d{4})\s+R\$\s*([\d\.,]+)',
        text
    )
    if m:
        r["ref_mes_ano"]  = f"{m.group(1)}/{m.group(2)}"
        r["vencimento"]   = _data(m.group(3))
        r["total_fatura"] = _num(m.group(4))

    # ── UC + datas de leitura + nr_dias ──────────────────────────────────────
    m = re.search(r'(\d{8})\s+(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})\s+(\d+)\n', text)
    if m:
        r["conta_contrato"]     = m.group(1)
        r["data_leitura_atual"] = _data(m.group(2))
        r["data_leitura_ant"]   = _data(m.group(3))
        r["nr_dias"]            = int(m.group(4))

    # ── Nota fiscal + data emissão ────────────────────────────────────────────
    m = re.search(
        r'NOTA FISCAL N[ºo]\s+(\d+).*?DATA DE EMISS[ÃA]O:\s*\n?(\d{2}/\d{2}/\d{4})',
        text, re.S
    )
    if m:
        r["nr_nota_fiscal"] = m.group(1)
        r["data_emissao"]   = _data(m.group(2))

    # ── Cliente ───────────────────────────────────────────────────────────────
    m = re.search(r'\n(JOSE[\w\s]+(?:DE\s+)?[\w]+)\nEST MUN', text)
    if m:
        r["cliente_nome"] = m.group(1).strip()

    # ── Subgrupo / modalidade ─────────────────────────────────────────────────
    m = re.search(r'Classifica[çc][aã]o:\s+Tarifa\s+(\w+)-(\w+)', text)
    if m:
        r["modalidade"] = m.group(1)   # Verde
        r["subgrupo"]   = m.group(2)   # A4

    # ── Nr medidor ────────────────────────────────────────────────────────────
    m = re.search(r'(\d{8})\s+Energia Ativa - kWh Ponta', text)
    if m:
        r["nr_medidor"] = m.group(1)

    # ── Demanda contratada ────────────────────────────────────────────────────
    m = re.search(r'Demanda kW\s+(\d+)', text)
    if m:
        r["demanda_contratada_kw"] = _num(m.group(1))

    # ── Consumo Ponta — TUSD ──────────────────────────────────────────────────
    m = re.search(
        r'Consumo Ponta \[KWh\] - TUSD.*?kWh\s+([\d\.,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d\.,]+)',
        text
    )
    if m:
        r["consumo_ponta_kwh"] = _num(m.group(1))
        r["tusd_ponta_sem"]    = _num(m.group(2))
        r["tusd_ponta_com"]    = _num(m.group(3))
        r["valor_tusd_ponta"]  = _num(m.group(4))

    # ── Consumo Fora Ponta — TUSD ─────────────────────────────────────────────
    m = re.search(
        r'Consumo Fora Ponta \[KWh\]-TUSD.*?kWh\s+([\d\.,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d\.,]+)',
        text
    )
    if m:
        r["consumo_fp_kwh"]  = _num(m.group(1))
        r["tusd_fp_sem"]     = _num(m.group(2))
        r["tusd_fp_com"]     = _num(m.group(3))
        r["valor_tusd_fp"]   = _num(m.group(4))

    # ── Consumo Ponta — TE ───────────────────────────────────────────────────
    m = re.search(
        r'Cons Ponta - TE.*?kWh\s+([\d\.,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d\.,]+)',
        text
    )
    if m:
        r["te_ponta_sem"]   = _num(m.group(2))
        r["te_ponta_com"]   = _num(m.group(3))
        r["valor_te_ponta"] = _num(m.group(4))

    # ── Consumo Fora Ponta — TE ──────────────────────────────────────────────
    m = re.search(
        r'Cons FPonta TE.*?kWh\s+([\d\.,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d\.,]+)',
        text
    )
    if m:
        r["te_fp_sem"]   = _num(m.group(2))
        r["te_fp_com"]   = _num(m.group(3))
        r["valor_te_fp"] = _num(m.group(4))

    # ── Demanda TUSD ─────────────────────────────────────────────────────────
    m = re.search(
        r'Demanda \[kW\] - TUSD.*?kW\s+([\d\.,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d\.,]+)',
        text
    )
    if m:
        r["demanda_medida_kw"] = _num(m.group(1))
        r["tusd_demanda_sem"]  = _num(m.group(2))
        r["tusd_demanda_com"]  = _num(m.group(3))
        r["valor_demanda"]     = _num(m.group(4))

    # ── Demanda Ultrapassagem ─────────────────────────────────────────────────
    m = re.search(
        r'Demanda Ultrap \[kW\] - TUSD.*?kW\s+([\d\.,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d\.,]+)',
        text
    )
    if m:
        r["demanda_ultrap_kw"]    = _num(m.group(1))
        r["tusd_ultrap_sem"]      = _num(m.group(2))
        r["tusd_ultrap_com"]      = _num(m.group(3))
        r["valor_demanda_ultrap"] = _num(m.group(4))

    # ── Energia Injetada FP — TUSD (GD) ──────────────────────────────────────
    m = re.search(
        r'Energia Atv Inj Fponta TUSD.*?kWh\s+([\d\.,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d\.,\-]+)',
        text
    )
    if m:
        r["injetada_fp_kwh"]   = _num(m.group(1))
        r["valor_inj_fp_tusd"] = _num(m.group(4).replace('-', '').strip())

    # ── Energia Injetada Ponta — TUSD (GD, quando existir) ───────────────────
    m = re.search(
        r'Energia Atv Inj Ponta TUSD.*?kWh\s+([\d\.,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d\.,\-]+)',
        text
    )
    if m:
        r["injetada_ponta_kwh"]   = _num(m.group(1))
        r["valor_inj_ponta_tusd"] = _num(m.group(4).replace('-', '').strip())

    # ── Energia Injetada FP — TE ─────────────────────────────────────────────
    m = re.search(
        r'Energia Atv Injetada Fponta TE.*?kWh\s+([\d\.,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d\.,\-]+)',
        text
    )
    if m:
        r["valor_inj_fp_te"] = _num(m.group(4).replace('-', '').strip())

    # ── Energia Injetada Ponta — TE ──────────────────────────────────────────
    m = re.search(
        r'Energia Atv Injetada Ponta TE.*?kWh\s+([\d\.,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d\.,\-]+)',
        text
    )
    if m:
        r["valor_inj_ponta_te"] = _num(m.group(4).replace('-', '').strip())

    # ── Bandeira tarifária ────────────────────────────────────────────────────
    m = re.search(r'Adicional Band (\w+) (?:Ponta|FPonta).*?kWh\s+([\d\.,]+)', text)
    if m:
        r["bandeira_cor"]   = m.group(1)
        r["valor_bandeira"] = _num(m.group(2))

    # ── CIP ───────────────────────────────────────────────────────────────────
    m = re.search(r'Contribui[çc][aã]o Custeio IP-CIP.*?([\d\.,]+)\n', text)
    if m:
        r["cosip"] = _num(m.group(1))

    # ── Tributos ──────────────────────────────────────────────────────────────
    m = re.search(r'ICMS\s+([\d\.,]+)\s+([\d,]+)\s+([\d\.,]+)', text)
    if m:
        r["icms_base"]  = _num(m.group(1))
        r["icms_aliq"]  = _num(m.group(2))
        r["icms_valor"] = _num(m.group(3))

    m = re.search(r'PIS/PASEP\s+([\d\.,]+)\s+([\d,]+)\s+([\d\.,]+)', text)
    if m:
        r["pis_base"]  = _num(m.group(1))
        r["pis_aliq"]  = _num(m.group(2))
        r["pis_valor"] = _num(m.group(3))

    m = re.search(r'COFINS\s+([\d\.,]+)\s+([\d,]+)\s+([\d\.,]+)', text)
    if m:
        r["cofins_base"]  = _num(m.group(1))
        r["cofins_aliq"]  = _num(m.group(2))
        r["cofins_valor"] = _num(m.group(3))

    # ── Total a Pagar ─────────────────────────────────────────────────────────
    m = re.search(r'Total a Pagar\s+([\d\.,]+)', text)
    if m:
        r["total_a_pagar"] = _num(m.group(1))

    # ── Taxa de perda técnica ─────────────────────────────────────────────────
    r["taxa_perda"] = bool(re.search(r'Taxa de perda', text, re.IGNORECASE))

    # ── Leituras do medidor ───────────────────────────────────────────────────
    # Formato: [medidor] [Tipo] - [unidade] [Posto] [lant] [latu] [mult] [cons]
    padrao_med = re.compile(
        r'\d{8}\s+(Energia Ativa|Energia Reativa|Demanda Ativa)\s*-\s*'
        r'(kWh|Kva|kVa|kVarh|kW)\s+(Ponta|Fora Ponta)'
        r'\s+(\d+)\s+(\d+)\s+([\d,]+)\s+([\d\.]+)',
        re.IGNORECASE
    )
    _MAP = {
        ("energia ativa",  "kwh",  "ponta"):       "med_kwh_ponta",
        ("energia ativa",  "kwh",  "fora ponta"):  "med_kwh_fp",
        ("demanda ativa",  "kw",   "ponta"):        "med_kw_ponta",
        ("demanda ativa",  "kw",   "fora ponta"):   "med_kw_fp",
        ("energia reativa","kva",  "ponta"):        "med_kvarh_ponta",
        ("energia reativa","kva",  "fora ponta"):   "med_kvarh_fp",
        ("energia reativa","kvarh","ponta"):        "med_kvarh_ponta",
        ("energia reativa","kvarh","fora ponta"):   "med_kvarh_fp",
    }
    for m in padrao_med.finditer(text):
        tipo  = m.group(1).lower().strip()
        unid  = m.group(2).lower().strip()
        posto = m.group(3).lower().strip()
        lant  = int(m.group(4))
        latu  = int(m.group(5))
        mult  = _num(m.group(6))
        cons  = _num(m.group(7))
        prefixo = _MAP.get((tipo, unid, posto))
        if prefixo:
            r[prefixo + "_lant"] = lant
            r[prefixo + "_latu"] = latu
            r[prefixo + "_mult"] = mult
            r[prefixo + "_cons"] = cons

    return r

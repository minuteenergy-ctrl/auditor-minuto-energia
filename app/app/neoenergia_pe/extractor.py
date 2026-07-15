# -*- coding: utf-8 -*-
"""
extractor.py - Parser unificado Neoenergia PE
Suporta:
  - Layout ANTIGO  (jul/21 - ago/22): NOTA FISCAL | FATURA
  - Layout DANFE   (set/22+):         DANFE - DOCUMENTO AUXILIAR
Autor: Minuto Energia
"""

import re
import pdfplumber
from pathlib import Path


def _num(s):
    if not s:
        return None
    s = str(s).strip().rstrip("-").replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _flatten(tbl):
    return [str(c) for row in tbl for c in row if c]


def _search(pattern, text, group=1, flags=re.IGNORECASE):
    m = re.search(pattern, text, flags)
    return m.group(group).strip() if m else None


# ─────────────────────────── LAYOUT ANTIGO ──────────────────────────────────

def _parse_antigo(text, tables):
    d = {}
    flat = _flatten(sum(tables, []) if tables else [[]])

    for cell in flat:
        if "DATA DE VENCIMENTO" in cell:
            m = re.search(r"DATA DE VENCIMENTO\s*\n?\s*(\d{2}/\d{2}/\d{4})", cell)
            if m: d["vencimento"] = m.group(1)
            m = re.search(r"TOTAL A PAGAR.*?\n?\s*([\d.,]+)", cell, re.IGNORECASE)
            if m: d["total_a_pagar"] = _num(m.group(1))
        if "CONTA CONTRATO" in cell:
            m = re.search(r"CONTA CONTRATO\s*\n?\s*(\d+)", cell)
            if m: d["conta_contrato"] = m.group(1)
            m = re.search(r"N[oO]\s+DO\s+CLIENTE\s*\n?\s*(\d+)", cell)
            if m: d["cod_cliente"] = m.group(1)
            m = re.search(r"N[oO]\s+DA\s+INSTALA..\s*\n?\s*(\d+)", cell)
            if m: d["cod_instalacao"] = m.group(1)
        if "EMISS" in cell.upper():
            m = re.search(r"EMISS..\s*\n?\s*(\d{2}/\d{2}/\d{4})", cell)
            if m: d["data_emissao"] = m.group(1)
            m = re.search(r"N.MERO DA NOTA FISCAL\s*\n?\s*(\d+)", cell)
            if m: d["nr_nota_fiscal"] = m.group(1)

    if "ref_mes_ano" not in d:
        m = re.search(r"(\d{2}/\d{4})\s+[\d.,]+\s+\d{2}/\d{2}/\d{4}", text)
        if m: d["ref_mes_ano"] = m.group(1)

    m = re.search(r"bandeira em vigor.*?(verde|amarela|vermelha|escassez|cinza)", text, re.IGNORECASE)
    if m:
        d["bandeira_cor"] = m.group(1).upper()
    else:
        m = re.search(r"Bandeira\s+(VERDE|AMARELA|VERMELHA|ESCASSEZ|CINZA)", text, re.IGNORECASE)
        if m: d["bandeira_cor"] = m.group(1).upper()

    m = re.search(r"Acr.scimo\s+Bandeira\s+\w+\s+([\d.,]+)", text)
    if m: d["valor_bandeira"] = _num(m.group(1))

    m = re.search(r"Contrib\.?\s+Ilum\.?\s+P.blica\s+Municipal\s+([\d.,]+)", text)
    if m: d["cosip"] = _num(m.group(1))

    m = re.search(r"ICMS Subven..o-CDE[^\n]+([\d.,]+)\s*$", text, re.MULTILINE)
    if m: d["icms_cde"] = _num(m.group(1))

    m = re.search(r"TOTAL DA FATURA\s+([\d.,]+)", text)
    if m: d["total_fatura"] = _num(m.group(1))

    for tbl in tables:
        for row in tbl:
            cells = [str(c) if c else "" for c in row]
            joined = "\n".join(cells)
            if "Consumo Ativo(kWh)-TUSD" not in joined:
                continue
            desc_cell = next((c for c in cells if "Consumo Ativo" in c), "")
            if "Bandeira" in desc_cell and "bandeira_cor" not in d:
                m = re.search(r"Acr.scimo Bandeira\s+(\w+)", desc_cell)
                if m: d["bandeira_cor"] = m.group(1).upper()
            for i, c in enumerate(cells):
                if re.search(r"\d+,\d{7}", c) and "consumo_kwh_tusd_qtd" not in d:
                    qtds = re.findall(r"(\d+,\d+)", c)
                    if qtds: d["consumo_kwh_tusd_qtd"] = _num(qtds[0])
                    if len(qtds) > 1: d["consumo_kwh_te_qtd"] = _num(qtds[1])
                if re.match(r"^0,\d{7,}\n0,\d{7,}\s*$", c.strip()) and "preco_tusd" not in d:
                    precos = re.findall(r"(0,\d+)", c)
                    if precos: d["preco_tusd"] = _num(precos[0])
                    if len(precos) > 1: d["preco_te"] = _num(precos[1])
                if re.match(r"^[\d]+,\d{2}\n[\d]+,\d{2}", c.strip()) and "valor_tusd" not in d:
                    vals = re.findall(r"([\d]+,\d{2})", c)
                    if len(vals) >= 2:
                        d["valor_tusd"] = _num(vals[0])
                        d["valor_te"]   = _num(vals[1])
                    if len(vals) > 2 and "cosip" not in d: d["cosip"] = _num(vals[2])
                    if len(vals) > 3 and "icms_cde" not in d: d["icms_cde"] = _num(vals[3])
            break

    m = re.search(r"Consumo Ativo\(kWh\)[ -]+TUSD\s+(0,\d+)", text)
    if m: d["tarifa_tusd_sem_trib"] = _num(m.group(1))
    m = re.search(r"Consumo Ativo\(kWh\)[ -]+TE\s+(0,\d+)", text)
    if m: d["tarifa_te_sem_trib"] = _num(m.group(1))

    for tbl in tables:
        for row in tbl:
            cells = [str(c).strip() if c else "" for c in row]
            nums = [_num(c) for c in cells if re.match(r"^[\d.,]+$", c)]
            if len(nums) >= 9:
                d.setdefault("icms_base",    nums[0])
                d.setdefault("icms_aliq",    nums[1])
                d.setdefault("icms_valor",   nums[2])
                d.setdefault("pis_base",     nums[3])
                d.setdefault("pis_aliq",     nums[4])
                d.setdefault("pis_valor",    nums[5])
                d.setdefault("cofins_base",  nums[6])
                d.setdefault("cofins_aliq",  nums[7])
                d.setdefault("cofins_valor", nums[8])
                break

    for tbl in tables:
        for row in tbl:
            cells = [str(c).strip() if c else "" for c in row]
            if any("CAT" in c for c in cells):
                for c in cells:
                    if re.match(r"^\d{7,10}$", c) and "nr_medidor" not in d:
                        d["nr_medidor"] = c
                    elif re.match(r"^\d{2}/\d{2}/\d{4}$", c):
                        if "data_leitura_anterior" not in d: d["data_leitura_anterior"] = c
                        elif "data_leitura_atual" not in d: d["data_leitura_atual"] = c
                    elif re.match(r"^[\d.]+,\d+$", c):
                        if "leitura_anterior" not in d: d["leitura_anterior"] = _num(c)
                        elif "leitura_atual" not in d: d["leitura_atual"] = _num(c)
                    elif re.match(r"^\d{1,3}$", c) and "nr_dias" not in d and int(c) > 0:
                        d["nr_dias"] = int(c)
                    elif re.match(r"^1,\d+$", c) and "constante_medidor" not in d:
                        d["constante_medidor"] = _num(c)

    m = re.search(r"PROXIMA LEITURA:\s*(\d{2}/\d{2}/\d{4})", text, re.IGNORECASE)
    if m: d["data_proxima_leitura"] = m.group(1)
    m = re.search(r"CLASSIFICA..O\s*\n?\s*(B\d[^-\n]+)", text, re.IGNORECASE)
    if m: d["classificacao"] = m.group(1).strip()
    if "tipo_fornecimento" not in d:
        m = re.search(r"((?:Bif.sico|Monof.sico|Trif.sico)\b[^,\n]*)", text, re.IGNORECASE)
        if m: d["tipo_fornecimento"] = m.group(1).strip()

    return d


# ─────────────────────────── LAYOUT DANFE ───────────────────────────────────

def _parse_danfe(text, tables, words=None):
    d = {}

    main_tbl_idx = None
    col_quant    = 4
    col_preco    = 7
    col_valor    = 8
    col_tarifa   = 22
    col_trib_lbl = 24
    col_trib_base= 26
    col_trib_alq = 27
    col_trib_val = 28

    for ti, tbl in enumerate(tables):
        for row in tbl:
            cells = [str(c) if c else "" for c in row]
            joined = " ".join(cells)
            if "ITENS DA FATURA" in joined and "QUANT." in joined:
                main_tbl_idx = ti
                for i, c in enumerate(cells):
                    cs = c.strip()
                    if "QUANT." in cs:
                        col_quant = i
                    elif "PRECO UNIT" in cs.upper() or "PRECO UNIT" in cs.upper():
                        col_preco = i
                    elif re.match(r"VALOR\s*\n?\s*\(R\$\)", cs) or re.match(r"\.\s*VALOR\s*\n", cs):
                        col_valor = i
                    elif "TARIFA" in cs and "UNIT" in cs:
                        col_tarifa = i
                    elif "T RIBUTO" in cs or cs.strip() == "TRIBUTO":
                        col_trib_lbl = i
                    elif "BASE DE" in cs and i > col_trib_lbl:
                        col_trib_base = i
                    elif "ALIQUOTA" in cs.upper() and "%" in cs and i > col_trib_lbl:
                        col_trib_alq = i
                    elif "VALOR" in cs and i > col_trib_lbl:
                        col_trib_val = i
                break
        if main_tbl_idx is not None:
            break

    main_tbl = tables[main_tbl_idx] if main_tbl_idx is not None else (tables[0] if tables else [])

    for tbl in tables:
        for row in tbl:
            for cell in row:
                if not cell: continue
                cs = str(cell)
                if "INSTALACAO" in cs.upper() or "INSTALA" in cs.upper():
                    m = re.search(r"(\d{5,7})", cs)
                    if m: d["cod_instalacao"] = m.group(1)
                if "DO CLIENTE" in cs.upper():
                    m = re.search(r"(\d{7,12})", cs)
                    if m:
                        d.setdefault("conta_contrato", m.group(1))
                        d.setdefault("cod_cliente",    m.group(1))

    if "conta_contrato" not in d:
        m = re.search(r"(\d{10})", text)
        if m: d.setdefault("conta_contrato", m.group(1))

    # cod_instalacao fallback via words (quando nao encontrado em tabelas)
    if "cod_instalacao" not in d and words:
        for w in words:
            if "INSTALA" in w["text"].upper() and w["x0"] > 200:
                ix0, itop = w["x0"], w["top"]
                digs = sorted(
                    [(ww["x0"], ww["text"]) for ww in words
                     if ix0 - 25 <= ww["x0"] <= ix0 + 20
                     and itop - 1 <= ww["top"] <= itop + 10
                     and re.match(r"^\d+$", ww["text"])],
                    key=lambda x: x[0]
                )
                code = "".join(t for _, t in digs)
                if re.match(r"^\d{5,7}$", code):
                    d["cod_instalacao"] = code
                break

    for row in main_tbl:
        cells = [str(c) if c else "" for c in row]
        joined = " ".join(cells)
        if "REF:M" in joined:
            for c in cells:
                if "REF:M" in c:
                    m = re.search(r"(\d{2}/\d{4})", c)
                    if m: d["ref_mes_ano"] = m.group(1)
                if "TOTAL A PAGAR" in c:
                    m = re.search(r"TOTAL A PAGAR\s*R?\$?\s*\n?\s*([\d.,]+)", c)
                    if m: d["total_a_pagar"] = _num(m.group(1))
                if "VENCIMENTO" in c and "DATA" not in c:
                    m = re.search(r"VENCIMENTO\s*\n?\s*(\d{2}/\d{2}/\d{4})", c)
                    if m: d["vencimento"] = m.group(1)
                if "EMISS" in c.upper():
                    m = re.search(r"(\d{2}/\d{2}/\d{4})", c)
                    if m: d["data_emissao"] = m.group(1)
        if "CLASSIFICA" in joined.upper():
            for c in cells:
                if "CLASSIFICA" in c.upper():
                    m = re.search(r"CLASSIFICA..O:\s*(.+?)(?:\n|$)", c, re.IGNORECASE)
                    if m: d["classificacao"] = m.group(1).strip()
        if "LEITURA ANTERIOR" in joined:
            for c in cells:
                if "LEITURA ANTERIOR" in c:
                    m = re.search(r"LEITURA ANTERIOR\s+(\d{2}/\d{2}/\d{4})", c)
                    if m: d["data_leitura_anterior"] = m.group(1)
                if "LEITURA ATUAL" in c:
                    m = re.search(r"LEITURA ATUAL\s+(\d{2}/\d{2}/\d{4})", c)
                    if m: d["data_leitura_atual"] = m.group(1)
                if "DE DIAS" in c.upper():
                    m = re.search(r"N.\s*DE\s*DIAS\s+(\d+)", c)
                    if m: d["nr_dias"] = int(m.group(1))
                if "PROXIMA LEITURA" in c.upper():
                    m = re.search(r"PROXIMA LEITURA\s+(\d{2}/\d{2}/\d{4})", c, re.IGNORECASE)
                    if m: d["data_proxima_leitura"] = m.group(1)

    for _tbl in tables:
        for _row in _tbl:
            for _cell in _row:
                if not _cell: continue
                cs = str(_cell)
                if "REF:M" in cs and "ref_mes_ano" not in d:
                    m = re.search(r"(\d{2}/\d{4})", cs)
                    if m: d["ref_mes_ano"] = m.group(1)
                if "VENCIMENTO" in cs and "DATA" not in cs and "vencimento" not in d:
                    m = re.search(r"(\d{2}/\d{2}/\d{4})", cs)
                    if m: d["vencimento"] = m.group(1)
                if "TOTAL A PAGAR" in cs and "total_a_pagar" not in d:
                    m = re.search(r"([\d.,]+)\s*$", cs.strip())
                    if m: d["total_a_pagar"] = _num(m.group(1))
                if "EMISS" in cs.upper() and "data_emissao" not in d:
                    m = re.search(r"(\d{2}/\d{2}/\d{4})", cs)
                    if m: d["data_emissao"] = m.group(1)

    if "ref_mes_ano" not in d:
        m = re.search(r"REF:M.S/ANO\s*\n?\s*(\d{2}/\d{4})", text)
        if m: d["ref_mes_ano"] = m.group(1)
    if "vencimento" not in d:
        m = re.search(r"VENCIMENTO\s*\n?\s*(\d{2}/\d{2}/\d{4})", text)
        if m: d["vencimento"] = m.group(1)
    if "total_a_pagar" not in d:
        m = re.search(r"TOTAL A PAGAR R?\$?\s*([\d.,]+)", text)
        if m: d["total_a_pagar"] = _num(m.group(1))
    if "data_emissao" not in d:
        m = re.search(r"DATA DE EMISS.O:\s*(\d{2}/\d{2}/\d{4})", text, re.IGNORECASE)
        if m: d["data_emissao"] = m.group(1)

    # ── itens da fatura (tabela) ──────────────────────────────────────────
    for row in main_tbl:
        cells = [str(c) if c else "" for c in row]
        if len(cells) <= col_quant: continue
        if "Consumo-TUSD" not in cells[0]: continue

        desc_cell  = cells[0]
        quant_cell = cells[col_quant] if col_quant < len(cells) else ""
        preco_cell = cells[col_preco] if col_preco < len(cells) else ""
        valor_cell = cells[col_valor] if col_valor < len(cells) else ""
        tarif_cell = cells[col_tarifa] if col_tarifa < len(cells) else ""

        qtds = re.findall(r"([\d]+,\d{2})", quant_cell)
        if qtds: d["consumo_kwh_tusd_qtd"] = _num(qtds[0])
        if len(qtds) > 1: d["consumo_kwh_te_qtd"] = _num(qtds[1])

        precos = re.findall(r"(0,\d{6,})", preco_cell)
        if precos: d["preco_tusd"] = _num(precos[0])
        if len(precos) > 1: d["preco_te"] = _num(precos[1])

        vals_pos, vals_neg = [], []
        for vm in re.finditer(r"([\d]+,\d{2})(-?)", valor_cell, re.MULTILINE):
            v = _num(vm.group(1))
            if vm.group(2) == "-": vals_neg.append(v)
            else: vals_pos.append(v)

        if vals_pos: d["valor_tusd"] = vals_pos[0]
        if len(vals_pos) > 1: d["valor_te"] = vals_pos[1]

        if "Bandeira" in desc_cell or "Band." in desc_cell:
            m = re.search(r"Acr.s\.?\s+Band(?:eira)?\.?\s+(\w+)", desc_cell)
            if m: d["bandeira_cor"] = m.group(1).upper()
            if len(vals_pos) > 2: d["valor_bandeira"] = vals_pos[2]
            if len(vals_pos) > 3: d["cosip"] = vals_pos[3]
            if len(vals_pos) > 4: d["icms_cde"] = vals_pos[4]
        else:
            if len(vals_pos) > 2: d["cosip"] = vals_pos[2]
            if len(vals_pos) > 3: d["icms_cde"] = vals_pos[3]

        tarifas = re.findall(r"(0,\d{6,})", tarif_cell)
        if tarifas: d["tarifa_tusd_sem_trib"] = _num(tarifas[0])
        if len(tarifas) > 1: d["tarifa_te_sem_trib"] = _num(tarifas[1])

        trib_lbl  = cells[col_trib_lbl]  if col_trib_lbl  < len(cells) else ""
        trib_base = cells[col_trib_base] if col_trib_base < len(cells) else ""
        trib_alq  = cells[col_trib_alq]  if col_trib_alq  < len(cells) else ""
        trib_val  = cells[col_trib_val]  if col_trib_val  < len(cells) else ""

        if "PIS" in trib_lbl:
            bases = re.findall(r"([\d.,]+)", trib_base)
            alqs  = re.findall(r"([\d.,]+)", trib_alq)
            vals  = re.findall(r"([\d.,]+)", trib_val)
            if len(bases) >= 3:
                d["pis_base"] = _num(bases[0]); d["cofins_base"] = _num(bases[1]); d["icms_base"] = _num(bases[2])
            if len(alqs) >= 3:
                d["pis_aliq"] = _num(alqs[0]);  d["cofins_aliq"] = _num(alqs[1]);  d["icms_aliq"] = _num(alqs[2])
            if len(vals) >= 3:
                d["pis_valor"] = _num(vals[0]); d["cofins_valor"] = _num(vals[1]); d["icms_valor"] = _num(vals[2])
        break

    # ── TUSD/TE via texto (SEMPRE sobrescreve) ────────────────────────────
    m = re.search(r"Consumo-TUSD\s+kWh\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)", text)
    if m:
        d["consumo_kwh_tusd_qtd"] = _num(m.group(1))
        d["preco_tusd"]           = _num(m.group(2))
        d["valor_tusd"]           = _num(m.group(3))
    m = re.search(r"Consumo-TE\s+kWh\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)", text)
    if m:
        d["consumo_kwh_te_qtd"] = _num(m.group(1))
        d["preco_te"]           = _num(m.group(2))
        d["valor_te"]           = _num(m.group(3))

    # ── Bandeira via texto (SEMPRE sobrescreve — col_valor pode sangrar) ──
    if "bandeira_cor" not in d:
        m = re.search(r"bandeira em vigor.*?(verde|amarela|vermelha|escassez|cinza)", text, re.IGNORECASE)
        if m: d["bandeira_cor"] = m.group(1).upper()
    m = re.search(r"Acr.s\.?\s+Band(?:eira)?\.?\s+\w+\s+([\d,]+)", text)
    d["valor_bandeira"] = _num(m.group(1)) if m else 0

    # ── COSIP / ICMS-CDE via texto (SEMPRE sobrescreve) ──────────────────
    # P.b. cobre tanto "Pub." (sem acento) quanto com acento no PDF
    m = re.search(r"Ilum\.?\s+P.b\.?\s+Municipal\s+([\d.,]+)", text)
    if m:
        d["cosip"] = _num(m.group(1))
    icms_cde_vals = re.findall(r"ICMS-CDE\s+\S+\s+([\d.,]+)", text)
    if icms_cde_vals:
        d["icms_cde"] = round(sum(_num(v) or 0 for v in icms_cde_vals), 2)

    # ── Parcelamento ──────────────────────────────────────────────────────
    if "valor_parcelamento" not in d:
        m = re.search(r"Parc\d+/\d+\s+\S+\s+([\d.,]+)", text)
        if m: d["valor_parcelamento"] = _num(m.group(1))

    # ── IPCA (correcao monetaria em parcelamentos) ────────────────────────
    ipca_vals = re.findall(r"IPCA-NF-\S+\s+([\d.,]+)", text)
    if ipca_vals:
        d["valor_ipca"] = round(sum(_num(v) or 0 for v in ipca_vals), 2)

    # ── SCEE / Compensacao — Imp.Som/Dim ─────────────────────────────────
    m = re.search(r"Imp\.Som/Dim-C/Impost\s+([\d,]+)-", text)
    if m:
        d["valor_imp_som_dim_c"] = -round(_num(m.group(1)) or 0, 2)
    m = re.search(r"Imp\.Som/Dim-S/Impost\s+([\d,]+)(-?)", text)
    if m:
        v = _num(m.group(1)) or 0
        d["valor_imp_som_dim_s"] = -v if m.group(2) == "-" else v

    # ── SCEE: kWh compensados (campo Informacoes Importantes) ────────────
    m = re.search(r"utilizados na unidade:\s*([\d,]+)\s*kWh", text, re.IGNORECASE)
    if m:
        d["scee_kwh_compensados"] = _num(m.group(1))
        d["is_scee"] = True

    # ── Religacao ─────────────────────────────────────────────────────────
    m = re.search(r"Relig\.?\s+U\.Consumidora\s+([\d.,]+)", text)
    if m:
        d["valor_religacao"] = _num(m.group(1))

    # ── Multas de NF (pode haver multiplas) ──────────────────────────────
    multa_vals = re.findall(r"Multa-NF\s+\S+\s+([\d.,]+)", text)
    if multa_vals:
        d["valor_multas_nf"] = round(sum(_num(v) or 0 for v in multa_vals), 2)

    # ── Juros de NF (pode haver multiplos) ───────────────────────────────
    juros_nf_vals = re.findall(r"Juros-NF\s+\S+\s+([\d.,]+)", text)
    if juros_nf_vals:
        d["valor_juros_nf"] = round(sum(_num(v) or 0 for v in juros_nf_vals), 2)

    # ── Encargos COSIP (JurosCOSIP, IPCACOSIP, Pla.JuroCOSIP) ───────────
    cosip_enc_vals = re.findall(r"(?:Juros|IPCA|Pla\.Juro)COSIP\s+([\d.,]+)", text)
    if cosip_enc_vals:
        d["valor_encargos_cosip"] = round(sum(_num(v) or 0 for v in cosip_enc_vals), 2)

    # ── Tributos via texto (SEMPRE sobrescreve) ───────────────────────────
    m = re.search(r"\bPIS\b\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)", text)
    if m:
        d["pis_base"] = _num(m.group(1)); d["pis_aliq"] = _num(m.group(2)); d["pis_valor"] = _num(m.group(3))
    m = re.search(r"\bCOFINS\b\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)", text)
    if m:
        d["cofins_base"] = _num(m.group(1)); d["cofins_aliq"] = _num(m.group(2)); d["cofins_valor"] = _num(m.group(3))
    m = re.search(r"\bICMS\b\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)", text)
    if m:
        d["icms_base"] = _num(m.group(1)); d["icms_aliq"] = _num(m.group(2)); d["icms_valor"] = _num(m.group(3))

    # ── total da fatura ───────────────────────────────────────────────────
    for row in main_tbl:
        cells = [str(c) if c else "" for c in row]
        if cells and cells[0].strip() == "TOTAL":
            if col_valor < len(cells) and cells[col_valor].strip():
                v = _num(cells[col_valor].strip())
                if v: d["total_fatura"] = v; break
            for c in cells[1:]:
                v = _num(c.strip())
                if v and v > 0: d["total_fatura"] = v; break
            break

    # ── medidor / leituras ────────────────────────────────────────────────
    for row in main_tbl:
        cells = [str(c) if c else "" for c in row]
        if not cells or not re.match(r"^\d{7,}", cells[0].strip()): continue
        # nr_medidor: apenas o primeiro (pode haver dois separados por \n)
        d["nr_medidor"] = cells[0].strip().split("\n")[0]
        reading_nums = []
        for c in cells[1:]:
            # celula pode conter multiplos medidores separados por \n — tomar 1a linha
            lines = [l.strip() for l in c.split("\n")]
            cs = lines[0]
            if re.match(r"^[\d.]+,\d+$", cs):
                reading_nums.append(_num(cs))
            elif "CONSUMO" in cs.upper():
                mm = re.search(r"([\d.,]+)", lines[-1] if len(lines) > 1 else cs)
                if mm: d.setdefault("consumo_medidor_kwh", _num(mm.group(1)))
        if len(reading_nums) >= 2:
            d["leitura_anterior"] = reading_nums[0]; d["leitura_atual"] = reading_nums[1]
        if len(reading_nums) >= 3: d["constante_medidor"] = reading_nums[2]
        if len(reading_nums) >= 4: d["consumo_medidor_kwh"] = reading_nums[3]
        break

    m = re.search(r"Protocolo de autoriza..o:\s*(\d+)", text, re.IGNORECASE)
    if m: d["nr_nota_fiscal"] = m.group(1)

    if "tipo_fornecimento" not in d:
        m = re.search(r"TIPO DE FORNECIMENTO[:\s]+(.+?)(?:\n|$)", text, re.IGNORECASE)
        if m: d["tipo_fornecimento"] = m.group(1).strip()

    return d


# ─────────────────────────── PARSER PRINCIPAL ───────────────────────────────

def parse_fatura(pdf_path):
    pdf_path = Path(pdf_path)
    result = {"arquivo": pdf_path.name, "layout": None, "erro": None}
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page   = pdf.pages[0]
            text   = page.extract_text() or ""
            tables = page.extract_tables()
            if "DANFE" in text:
                result["layout"] = "DANFE"
                words = page.extract_words()
                fields = _parse_danfe(text, tables, words)
            elif "NOTA FISCAL | FATURA" in text:
                result["layout"] = "ANTIGO"
                fields = _parse_antigo(text, tables)
            else:
                result["layout"] = "DESCONHECIDO"
                result["erro"]   = "Layout nao reconhecido"
                return result
            result.update(fields)
    except Exception as e:
        result["erro"] = str(e)
    return result

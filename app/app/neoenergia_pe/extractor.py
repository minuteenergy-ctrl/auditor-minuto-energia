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
            m = re.search(r"N.?\s+DO\s+CLIENTE\s*\n?\s*(\d+)", cell)
            if m: d["cod_cliente"] = m.group(1)
            # Nº DA INSTALAÇÃO ou No DA INSTALACAO (o º nao esta em [oO])
            m = re.search(r"INSTALA\S*\s*\n?\s*(\d{5,7})", cell)
            if m: d["cod_instalacao"] = m.group(1)
        if "EMISS" in cell.upper():
            # "EMISSÃO DA NOTA FISCAL\n18/03/2022" — ignora o texto entre o label e a data
            m = re.search(r"EMISS[^\n]*\n\s*(\d{2}/\d{2}/\d{4})", cell)
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
                    qtds = re.findall(r"([\d.]+,\d+)", c)
                    if qtds: d["consumo_kwh_tusd_qtd"] = _num(qtds[0])
                    if len(qtds) > 1: d["consumo_kwh_te_qtd"] = _num(qtds[1])
                if re.match(r"^0,\d{7,}\n0,\d{7,}\s*$", c.strip()) and "preco_tusd" not in d:
                    precos = re.findall(r"(0,\d+)", c)
                    if precos: d["preco_tusd"] = _num(precos[0])
                    if len(precos) > 1: d["preco_te"] = _num(precos[1])
                if re.match(r"^[\d.]+,\d{2}\n[\d.]+,\d{2}", c.strip()) and "valor_tusd" not in d:
                    vals = re.findall(r"([\d.]+,\d{2})", c)
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
                    elif re.match(r"VALOR\s*\n?\s*\(R\$\)", cs) or re.match(r"\.\s*VALOR\s*\n", cs):
                        col_valor = i
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

        qtds = re.findall(r"([\d.]+,\d{2})", quant_cell)
        if qtds: d["consumo_kwh_tusd_qtd"] = _num(qtds[0])
        if len(qtds) > 1: d["consumo_kwh_te_qtd"] = _num(qtds[1])

        precos = re.findall(r"(0,\d{6,})", preco_cell)
        if precos: d["preco_tusd"] = _num(precos[0])
        if len(precos) > 1: d["preco_te"] = _num(precos[1])

        vals_pos, vals_neg = [], []
        for vm in re.finditer(r"([\d.]+,\d{2})(-?)", valor_cell, re.MULTILINE):
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
    m = re.search(r"Consumo-TUSD\s+kWh\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)", text)
    if m:
        d["consumo_kwh_tusd_qtd"] = _num(m.group(1))
        d["preco_tusd"]           = _num(m.group(2))
        d["valor_tusd"]           = _num(m.group(3))
    m = re.search(r"Consumo-TE\s+kWh\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)", text)
    if m:
        d["consumo_kwh_te_qtd"] = _num(m.group(1))
        d["preco_te"]           = _num(m.group(2))
        d["valor_te"]           = _num(m.group(3))

    # ── Bandeira via texto (SEMPRE sobrescreve — col_valor pode sangrar) ──
    if "bandeira_cor" not in d:
        m = re.search(r"bandeira em vigor.*?(verde|amarela|vermelha|escassez|cinza)", text, re.IGNORECASE)
        if m: d["bandeira_cor"] = m.group(1).upper()
    # Captura TODAS as linhas de bandeira (pode haver 2 quando o ciclo cruza troca de bandeira)
    # Cobre: "Acrés. Band. AMARELA", "Acrés.Bd.VERMELHA-P2", "Acréscimo Bandeira VERDE", etc.
    band_vals = re.findall(r"Acr[eé]s\.?\s*(?:Band|Bd)\.?\s*\S+\s+([\d,]+)", text)
    d["valor_bandeira"] = round(sum(_num(v) or 0 for v in band_vals), 2) if band_vals else 0

    # ── COSIP / ICMS-CDE via texto (SEMPRE sobrescreve) ──────────────────
    # P.b. cobre tanto "Pub." (sem acento) quanto com acento no PDF
    icms_cde_vals = re.findall(r"ICMS-CDE\s+\S+\s+([\d.,]+)", text)
    m = re.search(r"Ilum\.?\s+P.b\.?\s+Municipal\s+([\d.,]+)", text)
    if m:
        d["cosip"] = _num(m.group(1))
        if not icms_cde_vals:
            # Sem ICMS-CDE: limpar qualquer icms_cde capturado erroneamente da tabela
            d.pop("icms_cde", None)
    elif icms_cde_vals:
        # Itens apos TUSD/TE sao ICMS-CDE, nao COSIP -- limpar valor capturado erroneamente da tabela
        d.pop("cosip", None)
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

    # ── Encargos COSIP (JurosCOSIP, IPCACOSIP, Pla.JuroCOSIP, Pla.IPCACOSIP) ──
    cosip_enc_vals = re.findall(r"(?:Juros|IPCA|Pla\.(?:Juro|IPCA))COSIP\s+([\d.,]+)", text)
    if cosip_enc_vals:
        d["valor_encargos_cosip"] = round(sum(_num(v) or 0 for v in cosip_enc_vals), 2)

    # ── Tributos via texto (SEMPRE sobrescreve) ───────────────────────────
    m = re.search(r"\bPIS\b\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)", text)
    if m:
        d["pis_base"] = _num(m.group(1)); d["pis_aliq"] = _num(m.group(2)); d["pis_valor"] = _num(m.group(3))
    m = re.search(r"\bCOFINS\b\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)", text)
    if m:
        d["cofins_base"] = _num(m.group(1)); d["cofins_aliq"] = _num(m.group(2)); d["cofins_valor"] = _num(m.group(3))
    m = re.search(r"\bICMS\b\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)", text)
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
        raw_med = cells[0].strip()
        d["nr_medidor"] = raw_med.split("\n")[0]
        dois_medidores = "\n" in raw_med
        if dois_medidores:
            d["nr_medidores"] = 2  # troca de medidor no ciclo
        reading_med1 = []  # medidor antigo (linhas[0] de cada celula)
        reading_med2 = []  # medidor novo  (linhas[1] de cada celula)
        for c in cells[1:]:
            parts = [l.strip() for l in c.split("\n")]
            c1 = parts[0] if parts else ""
            c2 = parts[1] if len(parts) > 1 else ""
            if re.match(r"^[\d.]+,\d+$", c1):
                reading_med1.append(_num(c1))
            if c2 and re.match(r"^[\d.]+,\d+$", c2):
                reading_med2.append(_num(c2))
            elif "CONSUMO" in c1.upper():
                mm = re.search(r"([\d.,]+)", parts[-1] if len(parts) > 1 else c1)
                if mm: d.setdefault("consumo_medidor_kwh", _num(mm.group(1)))
        if len(reading_med1) >= 2:
            d["leitura_anterior"] = reading_med1[0]; d["leitura_atual"] = reading_med1[1]
        if len(reading_med1) >= 3: d["constante_medidor"] = reading_med1[2]
        if len(reading_med1) >= 4: d["consumo_medidor_kwh"] = reading_med1[3]
        # segundo medidor (novo) — leituras auditaveis quando lant > 0
        if len(reading_med2) >= 2:
            d["leitura_anterior_2"] = reading_med2[0]
            d["leitura_atual_2"]    = reading_med2[1]
        break

    m = re.search(r"Protocolo de autoriza..o:\s*(\d+)", text, re.IGNORECASE)
    if m: d["nr_nota_fiscal"] = m.group(1)

    if "tipo_fornecimento" not in d:
        m = re.search(r"TIPO DE FORNECIMENTO[:\s]+(.+?)(?:\n|$)", text, re.IGNORECASE)
        if m: d["tipo_fornecimento"] = m.group(1).strip()

    return d


# ─────────────────────────── LAYOUT DANFE MT (Grupo A) ─────────────────────

def _parse_danfe_mt(text):
    """
    Extrai campos especificos de faturas MT (Grupo A) layout DANFE Neoenergia PE.
    Chamado apos _parse_danfe() com texto de TODAS as paginas concatenadas.
    Acrescenta campos MT e limpa campos garbled do _parse_danfe().
    """
    d = {"is_mt": True}
    # Anula constante_medidor garbled da secao MEDIDOR (colunas misalinhadas em MT DANFE)
    d["constante_medidor"] = None

    # Labels que encerram a secao de numericos de cada linha de item:
    _SUFIXO = r"\b(?:PIS|COFINS|ICMS|GRANDEZAS|Montante)\b"

    def _item_nums(pattern, unit):
        """
        Retorna lista de floats (virgula decimal) da linha de item, ANTES de qualquer
        label de tributo/grandeza que possa aparecer na mesma linha (artefato pdfplumber).
        """
        m = re.search(pattern + r"\s+" + unit + r"\s+([^\n]+)", text, re.IGNORECASE)
        if not m:
            return []
        linha = re.split(_SUFIXO, m.group(1))[0]
        return re.findall(r"[\d.]+,\d+", linha)

    def _mt_item(pattern, unit):
        """(qtd, preco_com, valor, tarifa_sem) de uma linha de item MT."""
        nums = _item_nums(pattern, unit)
        if len(nums) < 3:
            return None, None, None, None
        # Estrutura: qtd | preco_com | valor | [red_base | red_val | aliq% | icms_item_val] | tarifa_sem
        qtd        = _num(nums[0])
        preco_com  = _num(nums[1])
        valor      = _num(nums[2])
        tarifa_sem = _num(nums[-1]) if len(nums) >= 4 else None
        return qtd, preco_com, valor, tarifa_sem

    # ── Subgrupo / Modalidade ─────────────────────────────────────────
    m = re.search(r"Tarifa\s+(A\d\S*)\s*[-–]\s*\w+\s*[-–]\s*(\w+)", text, re.IGNORECASE)
    if m:
        d["subgrupo"]   = m.group(1).strip()
        d["modalidade"] = m.group(2).strip()

    # ── Grandezas Contratadas ─────────────────────────────────────────
    # Ex: "Montante de Uso Contratado FP 120"  (sem "kW" na mesma linha)
    for posto, key in [("FP", "dem_contratada_fp_kw"), ("NP", "dem_contratada_np_kw")]:
        m = re.search(
            r"Montante de Uso Contratado\s+" + posto + r"\s+([\d.,]+)",
            text, re.IGNORECASE
        )
        if m:
            d[key] = _num(m.group(1))

    # ── Itens: Uso Sistema Fio NP / FP ───────────────────────────────
    qtd, preco, valor, tarifa = _mt_item(r"Uso Sistema Fio NP", r"kW")
    d["dem_fio_np_kw"]      = qtd
    d["preco_fio_np_com"]   = preco
    d["valor_fio_np"]       = valor
    d["tarifa_fio_np_sem"]  = tarifa

    qtd, preco, valor, tarifa = _mt_item(r"Uso Sistema Fio FP", r"kW")
    d["dem_fio_fp_kw"]      = qtd
    d["preco_fio_fp_com"]   = preco
    d["valor_fio_fp"]       = valor
    d["tarifa_fio_fp_sem"]  = tarifa

    # ── Itens: Uso Sistema Encargo NP / FP ───────────────────────────
    qtd, preco, valor, tarifa = _mt_item(r"Uso Sistema Encar\.?\s*NP", r"kWh")
    d["consumo_encar_np_kwh"] = qtd
    d["preco_encar_np_com"]   = preco
    d["valor_encar_np"]       = valor
    d["tarifa_encar_np_sem"]  = tarifa

    qtd, preco, valor, tarifa = _mt_item(r"Uso Sistema Encar\.?\s*FP", r"kWh")
    d["consumo_encar_fp_kwh"] = qtd
    d["preco_encar_fp_com"]   = preco
    d["valor_encar_fp"]       = valor
    d["tarifa_encar_fp_sem"]  = tarifa

    # ── Itens: Demanda Reativa Excedente NP / FP ─────────────────────
    qtd, preco, valor, tarifa = _mt_item(r"Dem\.?\s*Reat[^\n]{0,20}NPonta", r"kVAr")
    d["dem_reat_np_kvar"]       = qtd
    d["valor_dem_reat_np"]      = valor
    d["tarifa_dem_reat_np_sem"] = tarifa

    # "Dem. Reativa Exc. FP" — descricao variavel; o unit "kVAr" identifica a linha
    nums_demreat_fp = _item_nums(r"Dem\.?\s*Reat(?:iva)?\s*Exc\.?\s*(?:FP\b|FPonta)", r"kVAr")
    if not nums_demreat_fp:
        # Fallback: segunda ocorrencia de uma linha com "Dem" + "Reat" + "kVAr"
        for m in re.finditer(r"Dem\.?\s*Reat[^\n]+kVAr\s+([^\n]+)", text, re.IGNORECASE):
            nums_tmp = re.findall(r"[\d.]+,\d+", re.split(_SUFIXO, m.group(1))[0])
            if len(nums_tmp) >= 1 and nums_demreat_fp == []:
                # Pula NP (primeiro match), pega FP (segundo)
                nums_demreat_fp = nums_tmp
                continue
            if nums_tmp:
                nums_demreat_fp = nums_tmp
                break
    d["dem_reat_fp_kvar"]       = _num(nums_demreat_fp[0]) if nums_demreat_fp else None
    d["valor_dem_reat_fp"]      = _num(nums_demreat_fp[2]) if len(nums_demreat_fp) >= 3 else None
    d["tarifa_dem_reat_fp_sem"] = _num(nums_demreat_fp[-1]) if len(nums_demreat_fp) >= 4 else None

    # ── Itens: Consumo Reativo Excedente NP / FP ─────────────────────
    qtd, preco, valor, _t = _mt_item(r"Cons\.?\s*Reat[^\n]{0,20}NPonta", r"kVARh")
    d["cons_reat_np_kvarh"] = qtd
    d["valor_cons_reat_np"] = valor

    qtd, preco, valor, _t = _mt_item(r"Cons\.?\s*Reat\s*[Ee]xc\.?[^\n]{0,15}FP(?:onta)?", r"kVARh")
    if qtd is None:
        # fallback sem "Exc": "Cons.Reat FPonta kVARh ..."
        qtd, preco, valor, _t = _mt_item(r"Cons\.?\s*Reat\s+(?:Exc\.?\s*)?FP(?:onta)?", r"kVARh")
    d["cons_reat_fp_kvarh"] = qtd
    d["valor_cons_reat_fp"] = valor

    # ── Demonstrativo: consumo e demanda medida (pagina 2) ────────────
    # Formato: "Consumo Ativo Na Ponta <lant> <latu> <const> <medido> <faturado> [grafico...]"
    # lant/latu sao inteiros (sem virgula); filtramos apenas floats com virgula.
    # Estrutura de floats: [lant_f, latu_f, const, medido, faturado]  (5 valores)
    for pattern, key_cons, key_cte in [
        (r"Consumo Ativo Na Ponta",       "consumo_ponta_kwh", "constante_medidor_np"),
        (r"Consumo Ativo Fora de Ponta",  "consumo_fp_kwh",    "constante_medidor_fp"),
    ]:
        m = re.search(pattern + r"\s+([^\n]+)", text, re.IGNORECASE)
        if m:
            # Truncar antes de dados de grafico (inteiros sem virgula, mes abrev)
            linha = re.split(r"\s+[A-Z]{3}\s", m.group(1))[0]
            floats = re.findall(r"[\d.]+,\d+", linha)
            if floats:
                d[key_cons] = _num(floats[-1])
            if len(floats) >= 3:
                d[key_cte] = _num(floats[-3])  # lant, latu, CONST, medido, fat

    for pattern, key_dem in [
        (r"Demanda M[áa]xima Na Ponta",             "demanda_medida_np_kw"),
        (r"Demanda M[áa]xima (?:Fora de Ponta|FP)", "demanda_medida_fp_kw"),
    ]:
        m = re.search(pattern + r"\s+([^\n]+)", text, re.IGNORECASE)
        if m:
            linha = re.split(r"\s+[A-Z]{3}\s|(?:Ponta|FP)\s*$", m.group(1))[0]
            floats = re.findall(r"[\d.]+,\d+", linha)
            if floats:
                d[key_dem] = _num(floats[-1])

    # ── Ajustes de valor (negativos, trailing dash) ────────────────────
    m = re.search(r"Dif\.?\s*Desc\.?\s*Ft\.?Alt-NP\s+([\d.,]+)-", text)
    if m:
        d["valor_dif_desc_np"] = -(_num(m.group(1)) or 0)
    m = re.search(r"Dif\.?\s*Desc\.?\s*Ft\.?Alt-FP\s+([\d.,]+)-", text)
    if m:
        d["valor_dif_desc_fp"] = -(_num(m.group(1)) or 0)

    # ── Imp.Som/Dim-C/Impost: para MT e positivo (ajuste tarifario) ──
    # Para BT com SCEE e negativo (trailing dash) e ja capturado em _parse_danfe.
    m = re.search(r"Imp\.Som/Dim-C/Impost\s+([\d.,]+)(?!\s*-)", text)
    if m:
        v = _num(m.group(1))
        if v is not None and v > 0:
            d["valor_imp_som_dim_mt"] = v

    # ── Desconto incondicional A4 (informativo, embutido no preco) ────
    m = re.search(r"Desconto Incondicional[^=\n]*=\s*R\$\s*([\d.,]+)", text, re.IGNORECASE)
    if m:
        d["desconto_a4_R$"] = _num(m.group(1))

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
                if "Uso Sistema Fio NP" in text:
                    # Fatura MT (Grupo A): Demonstrativo fica na pagina 2, ler todas as paginas
                    full_text = "\n".join(
                        (pdf.pages[i].extract_text() or "") for i in range(len(pdf.pages))
                    )
                    fields.update(_parse_danfe_mt(full_text))
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

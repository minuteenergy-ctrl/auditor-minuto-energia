# -*- coding: utf-8 -*-
"""
extractor.py - Parser unificado Neoenergia PE
Suporta:
  - Layout ANTIGO  (jul/21 – ago/22): "NOTA FISCAL | FATURA"
  - Layout DANFE   (set/22+):         "DANFE - DOCUMENTO AUXILIAR"
    • Sub-layout standard (22-25): cols VALOR=8
    • Sub-layout compact  (26+):   cols VALOR=6
Autor: Minuto Energia
"""

import re
import pdfplumber
from pathlib import Path


# ─────────────────────────── helpers ────────────────────────────────────────

def _num(s):
    """Converte string BR '1.234,56' → float. Retorna None se inválido."""
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

    # ── cabeçalho (leitura das células de tabela) ────────────────────────
    for cell in flat:
        if "DATA DE VENCIMENTO" in cell:
            m = re.search(r"DATA DE VENCIMENTO\s*\n?\s*(\d{2}/\d{2}/\d{4})", cell)
            if m:
                d["vencimento"] = m.group(1)
            m = re.search(r"TOTAL A PAGAR.*?\n?\s*([\d.,]+)", cell, re.IGNORECASE)
            if m:
                d["total_a_pagar"] = _num(m.group(1))

        if "CONTA CONTRATO" in cell:
            m = re.search(r"CONTA CONTRATO\s*\n?\s*(\d+)", cell)
            if m:
                d["conta_contrato"] = m.group(1)
            m = re.search(r"N[ºO]\s+DO\s+CLIENTE\s*\n?\s*(\d+)", cell)
            if m:
                d["cod_cliente"] = m.group(1)
            m = re.search(r"N[ºO]\s+DA\s+INSTALA[ÇC][AÃ]O\s*\n?\s*(\d+)", cell)
            if m:
                d["cod_instalacao"] = m.group(1)

        if "EMISSÃO" in cell or "EMISSAO" in cell:
            m = re.search(r"EMISS[AÃ]O.*?\n?\s*(\d{2}/\d{2}/\d{4})", cell)
            if m:
                d["data_emissao"] = m.group(1)
            m = re.search(r"N[ÚU]MERO DA NOTA FISCAL\s*\n?\s*(\d+)", cell)
            if m:
                d["nr_nota_fiscal"] = m.group(1)

    # ── ref mês/ano (rodapé) ─────────────────────────────────────────────
    if "ref_mes_ano" not in d:
        m = re.search(r"(\d{2}/\d{4})\s+[\d.,]+\s+\d{2}/\d{2}/\d{4}", text)
        if m:
            d["ref_mes_ano"] = m.group(1)

    # ── bandeira ─────────────────────────────────────────────────────────
    m = re.search(r"bandeira em vigor.*?(verde|amarela|vermelha|escassez|cinza)",
                  text, re.IGNORECASE)
    if m:
        d["bandeira_cor"] = m.group(1).upper()
    else:
        m = re.search(r"Bandeira\s+(VERDE|AMARELA|VERMELHA|ESCASSEZ|CINZA)", text, re.IGNORECASE)
        if m:
            d["bandeira_cor"] = m.group(1).upper()

    m = re.search(r"Acréscimo Bandeira\s+\w+\s+([\d.,]+)", text)
    if m:
        d["valor_bandeira"] = _num(m.group(1))

    # ── COSIP / ICMS-CDE ─────────────────────────────────────────────────
    m = re.search(r"Contrib\.?\s+Ilum\.?\s+P[úu]blica\s+Municipal\s+([\d.,]+)", text)
    if m:
        d["cosip"] = _num(m.group(1))

    m = re.search(r"ICMS Subven[çc][aã]o-CDE[^\n]+([\d.,]+)\s*$", text, re.MULTILINE)
    if m:
        d["icms_cde"] = _num(m.group(1))

    # ── total da fatura ───────────────────────────────────────────────────
    m = re.search(r"TOTAL DA FATURA\s+([\d.,]+)", text)
    if m:
        d["total_fatura"] = _num(m.group(1))

    # ── itens TUSD / TE (via tabela) ─────────────────────────────────────
    for tbl in tables:
        for row in tbl:
            cells = [str(c) if c else "" for c in row]
            joined = "\n".join(cells)
            if "Consumo Ativo(kWh)-TUSD" not in joined:
                continue

            # Célula de descrição
            desc_cell = next((c for c in cells if "Consumo Ativo" in c), "")

            # Bandeira dentro da descrição
            if "Bandeira" in desc_cell and "bandeira_cor" not in d:
                m = re.search(r"Acréscimo Bandeira\s+(\w+)", desc_cell)
                if m:
                    d["bandeira_cor"] = m.group(1).upper()

            # Quantidades: célula que contém "100,0000000" ou similar (7+ decimais)
            for i, c in enumerate(cells):
                if re.search(r"\d+,\d{7}", c) and "consumo_kwh_tusd_qtd" not in d:
                    qtds = re.findall(r"(\d+,\d+)", c)
                    if qtds:
                        d["consumo_kwh_tusd_qtd"] = _num(qtds[0])
                    if len(qtds) > 1:
                        d["consumo_kwh_te_qtd"] = _num(qtds[1])

                # Preços unitários (com trib): "0,50043497\n0,39775264"
                if re.match(r"^0,\d{7,}\n0,\d{7,}\s*$", c.strip()) and "preco_tusd" not in d:
                    precos = re.findall(r"(0,\d+)", c)
                    if precos:
                        d["preco_tusd"] = _num(precos[0])
                    if len(precos) > 1:
                        d["preco_te"] = _num(precos[1])

                # Valores dos itens: linha com 2+ valores "xx,xx"
                if re.match(r"^[\d]+,\d{2}\n[\d]+,\d{2}", c.strip()) and "valor_tusd" not in d:
                    vals = re.findall(r"([\d]+,\d{2})", c)
                    if len(vals) >= 2:
                        d["valor_tusd"] = _num(vals[0])
                        d["valor_te"]   = _num(vals[1])
                    if len(vals) > 2 and "cosip" not in d:
                        d["cosip"] = _num(vals[2])
                    if len(vals) > 3 and "icms_cde" not in d:
                        d["icms_cde"] = _num(vals[3])

            break  # primeira linha de Consumo Ativo basta

    # Tarifas sem trib (texto)
    m = re.search(r"Consumo Ativo\(kWh\)[ -]+TUSD\s+(0,\d+)", text)
    if m:
        d["tarifa_tusd_sem_trib"] = _num(m.group(1))
    m = re.search(r"Consumo Ativo\(kWh\)[ -]+TE\s+(0,\d+)", text)
    if m:
        d["tarifa_te_sem_trib"] = _num(m.group(1))

    # ── tributos (linha de valores numéricos: base%, valor x3) ───────────
    for tbl in tables:
        for row in tbl:
            cells = [str(c).strip() if c else "" for c in row]
            # Linha de dados: começa com número e tem pelo menos 9 campos numéricos
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

    # ── medidor / leituras ────────────────────────────────────────────────
    for tbl in tables:
        for row in tbl:
            cells = [str(c).strip() if c else "" for c in row]
            if any("CAT" in c for c in cells):
                for c in cells:
                    if re.match(r"^\d{7,10}$", c) and "nr_medidor" not in d:
                        d["nr_medidor"] = c
                    elif re.match(r"^\d{2}/\d{2}/\d{4}$", c):
                        if "data_leitura_anterior" not in d:
                            d["data_leitura_anterior"] = c
                        elif "data_leitura_atual" not in d:
                            d["data_leitura_atual"] = c
                    elif re.match(r"^[\d.]+,\d+$", c):
                        if "leitura_anterior" not in d:
                            d["leitura_anterior"] = _num(c)
                        elif "leitura_atual" not in d:
                            d["leitura_atual"] = _num(c)
                    elif re.match(r"^\d{1,3}$", c) and "nr_dias" not in d and int(c) > 0:
                        d["nr_dias"] = int(c)
                    elif re.match(r"^1,\d+$", c) and "constante_medidor" not in d:
                        d["constante_medidor"] = _num(c)

    # Próxima leitura
    m = re.search(r"DATA PREVISTA DA PR[ÓO]XIMA LEITURA:\s*(\d{2}/\d{2}/\d{4})", text)
    if m:
        d["data_proxima_leitura"] = m.group(1)

    # Classificação
    m = re.search(r"CLASSIFICA[ÇC][AÃ]O\s*\n?\s*(B\d[^-\n]+)", text)
    if m:
        d["classificacao"] = m.group(1).strip()

    return d


# ─────────────────────────── LAYOUT DANFE ───────────────────────────────────

def _parse_danfe(text, tables):
    d = {}

    # ── localiza a tabela principal e o índice da coluna VALOR ───────────
    main_tbl_idx = None
    col_quant    = 4   # default
    col_preco    = 7   # default (PRECO UNIT. COM TRIB.)
    col_valor    = 8   # default (VALOR R$)
    col_tarifa   = 22  # default (TARIFA UNIT sem trib)
    col_trib_lbl = 24  # default (PIS/COFINS/ICMS labels)
    col_trib_base= 26  # default
    col_trib_alq = 27  # default
    col_trib_val = 28  # default

    for ti, tbl in enumerate(tables):
        for row in tbl:
            cells = [str(c) if c else "" for c in row]
            joined = " ".join(cells)
            if "ITENS DA FATURA" in joined and "QUANT." in joined:
                main_tbl_idx = ti
                # Detecta índices dinamicamente
                for i, c in enumerate(cells):
                    cs = c.strip()
                    if "QUANT." in cs:
                        col_quant = i
                    elif "PREÇO UNIT." in cs or "PRECO UNIT." in cs:
                        col_preco = i
                    elif re.match(r"VALOR\s*\n\s*\(R\$\)", cs):
                        # Coluna de valor total dos itens: "VALOR\n(R$)"
                        col_valor = i
                    elif "TARIFA" in cs and "UNIT" in cs:
                        col_tarifa = i
                    elif "T RIBUTO" in cs or cs.strip() == "TRIBUTO":
                        col_trib_lbl = i
                    elif "BASE DE" in cs and i > col_trib_lbl:
                        col_trib_base = i
                    elif "ALÍQUOTA" in cs and "%" in cs and i > col_trib_lbl:
                        col_trib_alq = i
                    elif "VALOR" in cs and i > col_trib_lbl:
                        # Coluna de valor dos tributos: "VALOR (R$)"
                        col_trib_val = i
                break
        if main_tbl_idx is not None:
            break

    main_tbl = tables[main_tbl_idx] if main_tbl_idx is not None else (tables[0] if tables else [])

    # ── tabela de instalação (geralmente tabela separada ou dentro da main) ─
    for tbl in tables:
        for row in tbl:
            for cell in row:
                if not cell:
                    continue
                cs = str(cell)
                if "CÓDIGO DA INSTALAÇÃO" in cs or "CODIGO DA INSTALACAO" in cs:
                    m = re.search(r"(\d{5,7})", cs)
                    if m:
                        d["cod_instalacao"] = m.group(1)
                if "CÓDIGO DO CLIENTE" in cs or "CODIGO DO CLIENTE" in cs:
                    m = re.search(r"(\d{7,12})", cs)
                    if m:
                        d.setdefault("conta_contrato", m.group(1))
                        d.setdefault("cod_cliente",    m.group(1))

    # Fallback instalação via texto (garbled)
    if "cod_instalacao" not in d:
        # Extrai dígitos isolados entre "CÓDIGO DA INSTALAÇÃO" e "CANDEIAS" no texto
        m_start = text.find("CÓDIGO DA INSTALAÇÃO")
        m_end   = text.find("CANDEIAS/PRAZERES")
        if m_start >= 0 and m_end > m_start:
            section = text[m_start:m_end]
            solo_digits = re.findall(r"(?:^|\n)\s*(\d)\s*(?:\n|$)", section)
            if len(solo_digits) >= 5:
                d["cod_instalacao"] = "".join(solo_digits)

    # Fallback conta via texto
    if "conta_contrato" not in d:
        m = re.search(r"(\d{10})", text)
        if m:
            d.setdefault("conta_contrato", m.group(1))

    # ── cabeçalho principal ───────────────────────────────────────────────
    for row in main_tbl:
        cells = [str(c) if c else "" for c in row]
        joined = " ".join(cells)

        if "REF:MÊS/ANO" in joined or "REF:MES/ANO" in joined:
            for c in cells:
                if "REF:M" in c:
                    m = re.search(r"(\d{2}/\d{4})", c)
                    if m:
                        d["ref_mes_ano"] = m.group(1)
                if "TOTAL A PAGAR" in c:
                    m = re.search(r"TOTAL A PAGAR\s*R?\$?\s*\n?\s*([\d.,]+)", c)
                    if m:
                        d["total_a_pagar"] = _num(m.group(1))
                if "VENCIMENTO" in c and "DATA" not in c:
                    m = re.search(r"VENCIMENTO\s*\n?\s*(\d{2}/\d{2}/\d{4})", c)
                    if m:
                        d["vencimento"] = m.group(1)
                if "DATA DE EMISSÃO" in c or "DATA DE EMISSAO" in c:
                    m = re.search(r"(\d{2}/\d{2}/\d{4})", c)
                    if m:
                        d["data_emissao"] = m.group(1)

        if "CLASSIFICAÇÃO" in joined or "CLASSIFICACAO" in joined:
            for c in cells:
                if "CLASSIFICAÇÃO" in c or "CLASSIFICACAO" in c:
                    m = re.search(r"CLASSIFICA[ÇC][AÃ]O:\s*(.+?)(?:\n|$)", c)
                    if m:
                        d["classificacao"] = m.group(1).strip()

        if "LEITURA ANTERIOR" in joined:
            for c in cells:
                if "LEITURA ANTERIOR" in c:
                    m = re.search(r"LEITURA ANTERIOR\s+(\d{2}/\d{2}/\d{4})", c)
                    if m:
                        d["data_leitura_anterior"] = m.group(1)
                if "LEITURA ATUAL" in c:
                    m = re.search(r"LEITURA ATUAL\s+(\d{2}/\d{2}/\d{4})", c)
                    if m:
                        d["data_leitura_atual"] = m.group(1)
                if "N° DE DIAS" in c or "Nº DE DIAS" in c:
                    m = re.search(r"N[°º]\s*DE\s*DIAS\s+(\d+)", c)
                    if m:
                        d["nr_dias"] = int(m.group(1))
                if "PRÓXIMA LEITURA" in c or "PROXIMA LEITURA" in c:
                    m = re.search(r"PR[ÓO]XIMA LEITURA\s+(\d{2}/\d{2}/\d{4})", c)
                    if m:
                        d["data_proxima_leitura"] = m.group(1)

    # Fallbacks via texto
    if "ref_mes_ano" not in d:
        m = re.search(r"REF:M[EÊ]S/ANO\s*\n?\s*(\d{2}/\d{4})", text)
        if m:
            d["ref_mes_ano"] = m.group(1)
    if "vencimento" not in d:
        m = re.search(r"VENCIMENTO\s*\n?\s*(\d{2}/\d{2}/\d{4})", text)
        if m:
            d["vencimento"] = m.group(1)
    if "total_a_pagar" not in d:
        m = re.search(r"TOTAL A PAGAR R?\$?\s*([\d.,]+)", text)
        if m:
            d["total_a_pagar"] = _num(m.group(1))
    if "data_emissao" not in d:
        m = re.search(r"DATA DE EMISS[AÃ]O:\s*(\d{2}/\d{2}/\d{4})", text)
        if m:
            d["data_emissao"] = m.group(1)

    # ── itens da fatura ───────────────────────────────────────────────────
    for row in main_tbl:
        cells = [str(c) if c else "" for c in row]
        if len(cells) <= col_quant:
            continue
        if "Consumo-TUSD" not in cells[0]:
            continue

        desc_cell  = cells[0]
        quant_cell = cells[col_quant] if col_quant < len(cells) else ""
        preco_cell = cells[col_preco] if col_preco < len(cells) else ""
        valor_cell = cells[col_valor] if col_valor < len(cells) else ""
        tarif_cell = cells[col_tarifa] if col_tarifa < len(cells) else ""

        # Quantidades: "409,00\n409,00"
        qtds = re.findall(r"([\d]+,\d{2})", quant_cell)
        if qtds:
            d["consumo_kwh_tusd_qtd"] = _num(qtds[0])
        if len(qtds) > 1:
            d["consumo_kwh_te_qtd"] = _num(qtds[1])

        # Preço com trib: "0,56163449\n0,44770473"
        precos = re.findall(r"(0,\d{6,})", preco_cell)
        if precos:
            d["preco_tusd"] = _num(precos[0])
        if len(precos) > 1:
            d["preco_te"] = _num(precos[1])

        # Valores dos itens: primeiro=TUSD, segundo=TE, terceiro=Bandeira ou COSIP
        # Filtrar valores positivos (ignorar negativos que são créditos)
        vals_raw = re.findall(r"([\d]+,\d{2})(?:-)?", valor_cell)
        # Separar positivos dos negativos
        all_val_tokens = re.finditer(r"([\d.,]+)(-)?(?=\n|$| )", valor_cell)
        vals_pos = []
        vals_neg = []
        for vm in re.finditer(r"([\d]+,\d{2})(-?)", valor_cell, re.MULTILINE):
            v = _num(vm.group(1))
            if vm.group(2) == "-":
                vals_neg.append(v)
            else:
                vals_pos.append(v)

        if vals_pos:
            d["valor_tusd"] = vals_pos[0]
        if len(vals_pos) > 1:
            d["valor_te"] = vals_pos[1]

        # Bandeira
        if "Bandeira" in desc_cell or "Band." in desc_cell:
            m = re.search(r"(?:Acrés\.?|Acréscimo)\s+Band(?:eira)?\.?\s+(\w+)", desc_cell)
            if m:
                d["bandeira_cor"] = m.group(1).upper()
            if len(vals_pos) > 2:
                d["valor_bandeira"] = vals_pos[2]
                if len(vals_pos) > 3:
                    d["cosip"] = vals_pos[3]
                if len(vals_pos) > 4:
                    d["icms_cde"] = vals_pos[4]
        else:
            if len(vals_pos) > 2:
                d["cosip"] = vals_pos[2]
            if len(vals_pos) > 3:
                d["icms_cde"] = vals_pos[3]

        # Tarifa sem trib: "0,42538000\n0,33909000"
        tarifas = re.findall(r"(0,\d{6,})", tarif_cell)
        if tarifas:
            d["tarifa_tusd_sem_trib"] = _num(tarifas[0])
        if len(tarifas) > 1:
            d["tarifa_te_sem_trib"] = _num(tarifas[1])

        # Tributos consolidados (PIS/COFINS/ICMS)
        trib_lbl  = cells[col_trib_lbl]  if col_trib_lbl  < len(cells) else ""
        trib_base = cells[col_trib_base] if col_trib_base < len(cells) else ""
        trib_alq  = cells[col_trib_alq]  if col_trib_alq  < len(cells) else ""
        trib_val  = cells[col_trib_val]  if col_trib_val  < len(cells) else ""

        if "PIS" in trib_lbl:
            bases = re.findall(r"([\d.,]+)", trib_base)
            alqs  = re.findall(r"([\d.,]+)", trib_alq)
            vals  = re.findall(r"([\d.,]+)", trib_val)
            if len(bases) >= 3:
                d["pis_base"]    = _num(bases[0])
                d["cofins_base"] = _num(bases[1])
                d["icms_base"]   = _num(bases[2])
            if len(alqs) >= 3:
                d["pis_aliq"]    = _num(alqs[0])
                d["cofins_aliq"] = _num(alqs[1])
                d["icms_aliq"]   = _num(alqs[2])
            if len(vals) >= 3:
                d["pis_valor"]    = _num(vals[0])
                d["cofins_valor"] = _num(vals[1])
                d["icms_valor"]   = _num(vals[2])

        break  # primeira linha de itens basta

    # Bandeira via texto (fallback)
    if "bandeira_cor" not in d:
        m = re.search(r"bandeira em vigor.*?(verde|amarela|vermelha|escassez|cinza)",
                      text, re.IGNORECASE)
        if m:
            d["bandeira_cor"] = m.group(1).upper()

    # COSIP / ICMS-CDE fallback via texto
    if "cosip" not in d:
        m = re.search(r"Ilum\.?\s+P[úu]b\.?\s+Municipal\s+([\d.,]+)", text)
        if m:
            d["cosip"] = _num(m.group(1))
    if "icms_cde" not in d:
        m = re.search(r"ICMS-CDE\s+\S+\s+([\d.,]+)", text)
        if m:
            d["icms_cde"] = _num(m.group(1))

    # ── total da fatura ───────────────────────────────────────────────────
    for row in main_tbl:
        cells = [str(c) if c else "" for c in row]
        if cells and cells[0].strip() == "TOTAL":
            # Total está na coluna col_valor
            if col_valor < len(cells) and cells[col_valor].strip():
                v = _num(cells[col_valor].strip())
                if v:
                    d["total_fatura"] = v
                    break
            # Fallback: primeiro valor numérico positivo após "TOTAL"
            for c in cells[1:]:
                v = _num(c.strip())
                if v and v > 0:
                    d["total_fatura"] = v
                    break
            break

    # ── medidor / leituras ────────────────────────────────────────────────
    for row in main_tbl:
        cells = [str(c) if c else "" for c in row]
        if not cells or not re.match(r"^\d{7,}", cells[0].strip()):
            continue
        d["nr_medidor"] = cells[0].strip()
        # Leituras são valores numéricos em ordem
        reading_nums = []
        for c in cells[1:]:
            cs = c.strip()
            if re.match(r"^[\d.]+,\d+$", cs):
                reading_nums.append(_num(cs))
            elif "CONSUMO" in cs.upper():
                # Às vezes consumo está no mesmo cell que "CONSUMO\nkWh Cob\n409,00"
                m = re.search(r"([\d.,]+)$", cs, re.MULTILINE)
                if m:
                    d.setdefault("consumo_medidor_kwh", _num(m.group(1)))
        if len(reading_nums) >= 2:
            d["leitura_anterior"] = reading_nums[0]
            d["leitura_atual"]    = reading_nums[1]
        if len(reading_nums) >= 3:
            d["constante_medidor"] = reading_nums[2]
        if len(reading_nums) >= 4:
            d["consumo_medidor_kwh"] = reading_nums[3]
        break

    # Nr nota fiscal (protocolo)
    m = re.search(r"Protocolo de autoriza[çc][aã]o:\s*(\d+)", text)
    if m:
        d["nr_nota_fiscal"] = m.group(1)

    return d


# ─────────────────────────── PARSER PRINCIPAL ───────────────────────────────

def parse_fatura(pdf_path):
    """
    Processa um PDF de fatura Neoenergia PE.
    Retorna dict com todos os campos extraídos + metadados.
    """
    pdf_path = Path(pdf_path)
    result = {
        "arquivo":  pdf_path.name,
        "layout":   None,
        "erro":     None,
    }

    try:
        with pdfplumber.open(pdf_path) as pdf:
            page   = pdf.pages[0]
            text   = page.extract_text() or ""
            tables = page.extract_tables()

            if "DANFE" in text:
                result["layout"] = "DANFE"
                fields = _parse_danfe(text, tables)
            elif "NOTA FISCAL | FATURA" in text:
                result["layout"] = "ANTIGO"
                fields = _parse_antigo(text, tables)
            else:
                result["layout"] = "DESCONHECIDO"
                result["erro"]   = "Layout não reconhecido"
                return result

            result.update(fields)

    except Exception as e:
        result["erro"] = str(e)

    return result


# ─
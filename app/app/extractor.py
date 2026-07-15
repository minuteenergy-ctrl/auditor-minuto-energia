"""
Extrator de fatura CPFL Piratininga

Suporta dois formatos:
  - Formato 1 ("Nota Fiscal"):  faturas ate ~set/2025 (layout antigo)
  - Formato 2 ("DANF3E"):       faturas a partir de out/2025 (layout novo)
"""
import re
import datetime
from collections import defaultdict
import pdfplumber


def _to_number(s):
    if s is None:
        return None
    s = str(s).strip()
    if not s or s == "-":
        return None
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _to_date(s):
    if not s:
        return None
    try:
        return datetime.datetime.strptime(s.strip(), "%d/%m/%Y").date()
    except ValueError:
        return None


def _extract_horizontal_lines(pdf_path):
    lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words(use_text_flow=True, extra_attrs=["upright"])
            horiz = [w for w in words if w.get("upright", True)]
            rows = defaultdict(list)
            for w in horiz:
                row_key = round(w["top"] / 8)
                rows[row_key].append(w)
            for k in sorted(rows.keys()):
                row_words = sorted(rows[k], key=lambda x: x["x0"])
                line = " ".join(w["text"] for w in row_words)
                lines.append(line)
    return lines


def _extract_bandeira_dias(full_text):
    """
    Extrai {patamar: dias} das linhas de bandeira da fatura CPFL Piratininga.

    Linha TUSD (L18): patamar no INICIO do ciclo de leitura
    Linha TE   (L19): "Vermelha N Dias P_x" -- tres casos:
      1) TUSD == P_x (mesmo patamar): P_x teve N dias no inicio + M dias do Pat_b
      2) TUSD eh Vermelha P_y, P_y != P_x (mudanca de patamar, mesma cor):
           OLD -- P_x eh o patamar FINAL; TUSD (P_y) teve N dias no inicio
      3) TUSD eh cor diferente (ex: Amarela, Verde):
           P_x teve N dias no inicio (maio); Pat_b adiciona M dias da cor final
    Linha Band (L20+): "Adicional de Bandeira COLOR MES/ANO ... M Dias"
    """
    dias = {}

    # Patamar da linha TUSD (inicio do ciclo)
    tusd_patamar = None
    m_tusd = re.search(
        r"\bTUSD\s+\w+/\d+\b[\s\S]{0,300}?\b(VERMELHA\s+P([12])|AMARELA|VERDE)\b",
        full_text, re.IGNORECASE
    )
    if m_tusd:
        raw = m_tusd.group(1).upper().strip()
        if "VERMELHA" in raw and m_tusd.group(2):
            tusd_patamar = f"VERMELHA-P{m_tusd.group(2)}"
        else:
            tusd_patamar = raw  # "AMARELA" ou "VERDE"

    # Linha TE: "Vermelha N Dias P_x"
    pat_te = re.compile(r"\bVERMELHA\s+(\d{1,3})\s+[Dd]ias?\s+P([12])\b", re.IGNORECASE)
    m_te = pat_te.search(full_text)

    te_patamar = None
    first_key  = None

    if m_te:
        te_n_dias  = int(m_te.group(1))
        te_patamar = f"VERMELHA-P{m_te.group(2)}"

        if tusd_patamar == te_patamar:
            # Caso 1 -- mesmo patamar (ex: jul/dez: P1→P1 ou P2→P2)
            # te_n_dias = dias do periodo inicial (Vermelha)
            # pat_b Amarela/Verde traz os dias do periodo final (cor diferente)
            first_key = te_patamar
        elif (tusd_patamar and tusd_patamar.startswith("VERMELHA")
              and te_patamar.startswith("VERMELHA")):
            # Caso 2 -- mudanca de patamar dentro da Vermelha (ex: ago: P1→P2, out: P2→P1)
            # TUSD = patamar inicial; te_patamar = patamar final
            # te_n_dias = dias do patamar inicial (TUSD)
            first_key = tusd_patamar
        else:
            # Caso 3 -- mudanca de cor (ex: jun: Amarela→Vermelha P1)
            # TUSD = cor inicial (ex: Amarela); te_patamar = cor final (ex: Vermelha P1)
            # te_n_dias = dias do periodo inicial (TUSD = Amarela)
            # pat_b linha Amarela traz "M Dias" = dias do periodo Vermelha (final)
            first_key = tusd_patamar  # ex: AMARELA

        dias[first_key] = te_n_dias

    # Indicador de Caso 3 para uso no pat_b abaixo
    is_case3 = (
        m_te is not None and
        tusd_patamar is not None and
        te_patamar is not None and
        not tusd_patamar.startswith("VERMELHA") and
        te_patamar.startswith("VERMELHA")
    )

    # Linha(s) "Adicional de Bandeira COLOR MES/ANO ... M Dias" -- periodo complementar
    pat_b = re.compile(
        r"Adicional\s+de\s+Bandeira\s+(VERMELHA|AMARELA|VERDE)\s+\w+/\d+[\s\S]{0,200}?(\d{1,3})\s+[Dd]ias\b",
        re.IGNORECASE
    )
    for bm in pat_b.finditer(full_text):
        cor    = bm.group(1).upper()
        m_dias = int(bm.group(2))

        if cor == "VERMELHA":
            if is_case3:
                # Caso 3: linha Adicional Vermelha tem "N Dias" = dias do periodo
                # nao-Vermelha (ja capturado via te_n_dias). Ignorar.
                continue
            elif tusd_patamar == te_patamar and te_patamar:
                # Caso 1: mesmo patamar -> final = mesmo
                final_key = te_patamar
            elif tusd_patamar and tusd_patamar.startswith("VERMELHA") and te_patamar:
                # Caso 2: mudanca de patamar -> final = te_patamar
                final_key = te_patamar
            elif tusd_patamar and tusd_patamar.startswith("VERMELHA"):
                final_key = tusd_patamar
            else:
                final_key = "VERMELHA-P1"
        else:
            if is_case3:
                # Caso 3: linha Adicional Amarela/Verde tem "M Dias" = dias do
                # periodo Vermelha (te_patamar), nao da cor atual
                final_key = te_patamar
            else:
                final_key = cor  # "AMARELA" ou "VERDE"

        dias[final_key] = dias.get(final_key, 0) + m_dias

    # Fallback legado (sem linha TE e sem linha Adicional com Dias)
    if not dias:
        pat_c = re.compile(
            r"\b(VERDE|AMARELA|VERMELHA)\s*(?:(P[12]))?\s+(\d{1,3})\s+[Dd]ias?\b",
            re.IGNORECASE
        )
        for m in pat_c.finditer(full_text):
            cor     = m.group(1).upper()
            pat_tag = m.group(2)
            n_dias  = int(m.group(3))
            if cor == "VERDE":
                key = "VERDE"
            elif cor == "AMARELA":
                key = "AMARELA"
            else:
                key = "VERMELHA-P" + (pat_tag[-1] if pat_tag else "1")
            dias[key] = dias.get(key, 0) + n_dias

    return dias


def _detect_bandeira_vigente(full_text, dias_patamar):
    if dias_patamar.get("VERMELHA-P2", 0) > 0:
        return "VERMELHA-P2"
    if dias_patamar.get("VERMELHA-P1", 0) > 0:
        return "VERMELHA-P1"
    if dias_patamar.get("AMARELA", 0) > 0:
        return "AMARELA"
    if re.search(r"\bVERMELHA\b.*?\bP2\b", full_text, re.IGNORECASE | re.DOTALL):
        return "VERMELHA-P2"
    if re.search(r"\bVERMELHA\b.*?\bP1\b", full_text, re.IGNORECASE | re.DOTALL):
        return "VERMELHA-P1"
    if re.search(r"\bVERMELHA\b", full_text, re.IGNORECASE):
        return "VERMELHA-P1"
    if re.search(r"\bAMARELA\b", full_text, re.IGNORECASE):
        return "AMARELA"
    return "VERDE"


# ---------------------------------------------------------------------------
# Formato 1 — "Nota Fiscal" (mai-set/2025)
# ---------------------------------------------------------------------------

def _extract_formato1(lines, full_text):
    data = {}

    m = re.search(r"N[oOº°]\s+(\d+)\s+S[eé]rie\s+C", full_text)
    if m:
        data["nota_fiscal"] = m.group(1)

    m = re.search(r"Emiss[ãa]o:\s*(\d{2}/\d{2}/\d{4})", full_text)
    if m:
        data["data_emissao"] = _to_date(m.group(1))

    m = re.search(r"Conta\s+Contrato\s+N[°º]\s*(\d+)", full_text)
    if m:
        data["conta_contrato"] = m.group(1)

    m = re.search(r"[Ll]eitura\s+[Pp]r[oó]ximo\s+M[eê]s[:\s]*(\d{2}/\d{2}/\d{4})", full_text)
    if m:
        data["proxima_leitura"] = _to_date(m.group(1))

    # www.cpfl.com.br {UC} {CodigoCliente} {MES/ANO} {vencimento} {total}
    m = re.search(
        r"www\.cpfl\.com\.br\s+(\d{5,12})\s+(\d{8,13})\s+([A-Z]{3}/\d{4})\s+(\d{2}/\d{2}/\d{4})\s+([\d.,]+)",
        full_text
    )
    if m:
        data["uc"] = m.group(1)
        data["codigo_cliente"] = m.group(2)
        data["mes_ref"] = m.group(3)
        data["data_vencimento"] = _to_date(m.group(4))
        data["total_a_pagar"] = _to_number(m.group(5))

    # Cliente + CPF + Classificacao
    m = re.search(
        r"([A-Z\xC0-\xDC][A-Z\xC0-\xDC\s]{4,40}?)\s+CPF:\s*([\d.*/-]+)\s+CLASSIFICA",
        full_text
    )
    if m:
        data["cliente_nome"] = m.group(1).strip()
        data["cliente_cpf"] = m.group(2).strip()

    m = re.search(r"CLASSIFICA[ÇC][ÃA]O:\s*([\w\s\-/.]+?)(?:\s+-\s+Bif|\s+-\s+Monof|\s+-\s+Trif|$)", full_text)
    if m:
        class_full = m.group(1).strip()
        data["classificacao"] = class_full
        m2 = re.search(r"\b(B[1-4][AaBb]?|A[1-4])\b", class_full)
        if m2:
            data["subgrupo"] = m2.group(1).upper()
        if "onvencional" in class_full:
            data["modalidade"] = "Convencional"
        elif "ranca" in class_full:
            data["modalidade"] = "Branca"

    # Tipo de fornecimento (captura variante completa, ex: "Bifásico a 3 condutores")
    m = re.search(
        r"((?:Bif[aá]sico|Monof[aá]sico|Trif[aá]sico)"
        r"(?:\s+a\s+(?:dois|três|2|3)\s+condutores)?)",
        full_text, re.IGNORECASE
    )
    if m:
        data["tipo_fornecimento"] = m.group(1).strip()

    # Aliquotas PIS/COFINS do cabecalho da tabela
    m = re.search(r"PIS/COFINS\s+([\d,]+)%\s+([\d,]+)%", full_text)
    if m:
        data["_pis_alq_pct"] = _to_number(m.group(1))
        data["_cofins_alq_pct"] = _to_number(m.group(2))

    # Itens
    items = []

    # TUSD
    m = re.search(
        r"0605Consumo Uso Sistema \[KWh\]-TUSD\s+\w+/\d+\s+([\d.,]+),?\d*\s+kWh\s+"
        r"([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)",
        full_text
    )
    if m:
        items.append({
            "tipo": "consumo_tusd",
            "descricao": "Consumo Uso Sistema [KWh]-TUSD",
            "unidade": "kWh",
            "quantidade": _to_number(m.group(1)),
            "preco_unit_com_trib": _to_number(m.group(2)),
            "valor": _to_number(m.group(3)),
            "base_icms": _to_number(m.group(4)),
            "aliq_icms": _to_number(m.group(5)),
            "icms": _to_number(m.group(6)),
            "pis": _to_number(m.group(7)),
            "cofins": _to_number(m.group(8)),
        })

    m = re.search(
        r"0601Consumo - TE\s+\w+/\d+\s+([\d.,]+),?\d*\s+kWh\s+"
        r"([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)",
        full_text
    )
    if m:
        items.append({
            "tipo": "consumo_te",
            "descricao": "Consumo - TE",
            "unidade": "kWh",
            "quantidade": _to_number(m.group(1)),
            "preco_unit_com_trib": _to_number(m.group(2)),
            "valor": _to_number(m.group(3)),
            "base_icms": _to_number(m.group(4)),
            "aliq_icms": _to_number(m.group(5)),
            "icms": _to_number(m.group(6)),
            "pis": _to_number(m.group(7)),
            "cofins": _to_number(m.group(8)),
        })

    band_pat = re.compile(
        # Aceita texto livre entre "Adicional" e "de Bandeira" (ex: quando
        # pdfplumber mescla a linha com sub-itens como "0804Juros de Mora").
        # Aceita tambem dois meses (ex: MAR/25 JUN/25) antes dos valores.
        r"0601Adicional[^\n]*?de\s+Bandeira\s+(VERMELHA(?:\s+P[12])?|AMARELA|VERDE)\s+(?:\w+/\d+\s+)+"
        r"([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)",
        re.IGNORECASE
    )
    for bm in band_pat.finditer(full_text):
        cor = re.sub(r"\s+", "-", bm.group(1).upper().strip())
        if cor == "VERMELHA":
            cor = "VERMELHA-P1"
        items.append({
            "tipo": "bandeira",
            "descricao": f"Adicional de Bandeira {cor}",
            "patamar": cor,
            "valor": _to_number(bm.group(2)),
            "base_icms": _to_number(bm.group(3)),
            "aliq_icms": _to_number(bm.group(4)),
            "icms": _to_number(bm.group(5)),
            "pis": _to_number(bm.group(6)),
            "cofins": _to_number(bm.group(7)),
        })

    m = re.search(r"[Bb][oô]nus.*?Lei.*?(\d{1,5}[.,]\d{2})", full_text)
    if m:
        items.append({"tipo": "bonus_itaipu", "descricao": "Bonus Itaipu / Lei 10438", "valor": _to_number(m.group(1))})

    data["itens"] = items

    # Tributos — Total Consolidado
    m = re.search(
        r"Total\s+Consolidado[\s\S]{0,20}?([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)",
        full_text
    )
    if not m:
        m = re.search(
            r"\n([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*\n",
            full_text
        )
    tributos = {}
    if m:
        total_pagar = _to_number(m.group(1))
        total_dist = _to_number(m.group(2))
        icms_val = _to_number(m.group(3))
        base_pis = _to_number(m.group(4))
        pis_val = _to_number(m.group(5))
        cofins_val = _to_number(m.group(6))
        pis_alq = data.pop("_pis_alq_pct", None) or 0
        cofins_alq = data.pop("_cofins_alq_pct", None) or 0
        tributos["icms"] = {"base": total_dist, "aliquota_pct": 18.0, "valor": icms_val}
        tributos["pis"] = {"base": base_pis, "aliquota_pct": pis_alq, "valor": pis_val}
        tributos["cofins"] = {"base": base_pis, "aliquota_pct": cofins_alq, "valor": cofins_val}
        if not data.get("total_a_pagar"):
            data["total_a_pagar"] = total_pagar
        data["total_fatura"] = total_dist
    data["tributos"] = tributos

    # Medidor — datas podem vir coladas: "20/05/202517/04/2025" (sem espaco)
    m = re.search(
        r"Consumo\s+kWh\s+([\d.,]+)\s+([\d.,]+)\s+(\d{8,10})\s+Ativa\s+(\d+)\s+"
        r"(\d{2}/\d{2}/\d{4})\s*(\d{2}/\d{2}/\d{4})\s+(\d+)\s+([\d.,]+)\s+Multipl\.\s+(\d+)",
        full_text
    )
    if m:
        data["medidores"] = [{
            "numero": m.group(3),
            "grandeza": "Energia Ativa",
            "leitura_anterior": _to_number(m.group(7)),
            "leitura_atual": _to_number(m.group(4)),
            "data_leitura_atual": _to_date(m.group(5)),
            "data_leitura_anterior": _to_date(m.group(6)),
            "constante": _to_number(m.group(8)),
            "consumo": _to_number(m.group(9)),
        }]
        data["leitura_atual"] = _to_date(m.group(5))
        data["leitura_anterior"] = _to_date(m.group(6))
        if data.get("leitura_atual") and data.get("leitura_anterior"):
            data["dias_ciclo"] = (data["leitura_atual"] - data["leitura_anterior"]).days
        if not data.get("uc"):
            data["uc"] = m.group(3)

    return data


# ---------------------------------------------------------------------------
# Formato 2 — DANF3E (out/2025 em diante)
# ---------------------------------------------------------------------------

def _extract_formato2(lines, full_text):
    data = {}

    # NF + data emissao
    m = re.search(
        r"NOTA FISCAL\s+N[OoºPp°]\s*(\d+)\s*[-/]?\s*S[ÉE]RIE\s*\d+\s*/?\s*DATA DE EMISS[ÃA]O:\s*\n?\s*(\d{2}/\d{2}/\d{4})",
        full_text
    )
    if m:
        data["nota_fiscal"] = m.group(1)
        data["data_emissao"] = _to_date(m.group(2))

    # Cabecalho: lote/roteiro/medidor/pag/datas
    m = re.search(
        r"\d+\s+SORBU\w+-\w+\s+(\d{8,10})\s+\d+/\d+\s+"
        r"(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})",
        full_text
    )
    if m:
        data["_medidor_id"] = m.group(1)
        if not data.get("data_emissao"):
            data["data_emissao"] = _to_date(m.group(2))
        data["proxima_leitura"] = _to_date(m.group(3))
        data["data_vencimento"] = _to_date(m.group(4))

    # Datas de leitura + dias — evita capturar o "24" de "24/11/2025"
    m = re.search(r"(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})\s+(\d{2,3})(?!/)", full_text)
    if m:
        data["leitura_atual"] = _to_date(m.group(1))
        data["leitura_anterior"] = _to_date(m.group(2))
        data["dias_ciclo"] = int(m.group(3))

    # Cliente nome + codigo
    for line in lines:
        line = line.strip()
        m = re.match(r"^([A-Z\xC0-\xDC][A-Z\xC0-\xDC\s]{5,40}?)\s+(\d{10,13})$", line)
        if m:
            data["cliente_nome"] = m.group(1).strip()
            data["codigo_cliente"] = m.group(2).strip()
            break

    # CPF
    m = re.search(r"CPF:\s*([\d.*/-]+)", full_text)
    if m:
        data["cliente_cpf"] = m.group(1).strip()

    # Endereco
    for i, line in enumerate(lines):
        if re.match(r"^[A-Z]{1,3}\s+[A-Z]", line.strip()) and "," in line:
            end_parts = [line.strip()]
            for j in range(i+1, min(i+3, len(lines))):
                nxt = lines[j].strip()
                if re.match(r"^\d{5}-\d{3}", nxt) or "SP" in nxt:
                    end_parts.append(nxt)
                    break
            if len(end_parts) > 1:
                data["cliente_endereco"] = " - ".join(end_parts)
                break

    # Classificacao — usa re.search para encontrar B1/B2/B3 dentro da string
    m = re.search(r"Classifica[cç][aã]o:\s*([\w\s]+?)\s+Tipo de Fornecimento:", full_text)
    if m:
        class_full = m.group(1).strip()
        data["classificacao"] = class_full
        m2 = re.search(r"\b(B[1-4][AaBb]?|A[1-4])\b", class_full)
        if m2:
            data["subgrupo"] = m2.group(1).upper()
        if "onvencional" in class_full:
            data["modalidade"] = "Convencional"
        elif "ranca" in class_full:
            data["modalidade"] = "Branca"

    # Tipo de fornecimento (captura variante completa, ex: "Bifásico a 3 condutores")
    m = re.search(
        r"((?:Bif[aá]sico|Monof[aá]sico|Trif[aá]sico)"
        r"(?:\s+a\s+(?:dois|três|2|3)\s+condutores)?)",
        full_text, re.IGNORECASE
    )
    if m:
        data["tipo_fornecimento"] = m.group(1).strip()

    # Mes ref + vencimento + total
    m = re.search(r"([A-Z]{3}/\d{4})\s+(\d{2}/\d{2}/\d{4})\s+R\$\s*([\d.,]+)", full_text)
    if m:
        data["mes_ref"] = m.group(1)
        if not data.get("data_vencimento"):
            data["data_vencimento"] = _to_date(m.group(2))
        data["total_a_pagar"] = _to_number(m.group(3))

    # Bloco tributos
    m1 = re.search(r"([\d,]+)%\s+([\d,]+)%\s+ICMS\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)", full_text)
    m2t = re.search(r"PIS/PASEP\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)", full_text)
    m3 = re.search(r"COFINS\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)", full_text)
    tributos = {}
    if m1:
        tributos["icms"] = {
            "base": _to_number(m1.group(3)),
            "aliquota_pct": _to_number(m1.group(4)),
            "valor": _to_number(m1.group(5)),
        }
    if m2t:
        tributos["pis"] = {
            "base": _to_number(m2t.group(1)),
            "aliquota_pct": _to_number(m2t.group(2)),
            "valor": _to_number(m2t.group(3)),
        }
    if m3:
        tributos["cofins"] = {
            "base": _to_number(m3.group(1)),
            "aliquota_pct": _to_number(m3.group(2)),
            "valor": _to_number(m3.group(3)),
        }
    data["tributos"] = tributos

    # Itens
    items = []

    m = re.search(
        r"Consumo Uso Sistema \[KWh\]-TUSD\s+\w+/\d+\s+kWh\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)",
        full_text
    )
    if m:
        items.append({
            "tipo": "consumo_tusd",
            "descricao": "Consumo Uso Sistema [KWh]-TUSD",
            "unidade": "kWh",
            "quantidade": _to_number(m.group(1)),
            "tarifa_unit_sem": _to_number(m.group(2)),
            "preco_unit_com_trib": _to_number(m.group(3)),
            "valor": _to_number(m.group(4)),
        })

    m = re.search(
        r"Consumo - TE\s+\w+/\d+\s+kWh\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)",
        full_text
    )
    if m:
        items.append({
            "tipo": "consumo_te",
            "descricao": "Consumo - TE",
            "unidade": "kWh",
            "quantidade": _to_number(m.group(1)),
            "tarifa_unit_sem": _to_number(m.group(2)),
            "preco_unit_com_trib": _to_number(m.group(3)),
            "valor": _to_number(m.group(4)),
        })

    band_pat = re.compile(
        r"Adicional de Bandeira\s+(VERMELHA(?:\s+P[12])?|AMARELA|VERDE)\s+\w+/\d+\s+kWh\s+([\d.,]+)",
        re.IGNORECASE
    )
    for bm in band_pat.finditer(full_text):
        cor = re.sub(r"\s+", "-", bm.group(1).upper().strip())
        if cor == "VERMELHA":
            cor = "VERMELHA-P1"
        items.append({
            "tipo": "bandeira",
            "descricao": f"Adicional de Bandeira {cor}",
            "patamar": cor,
            "valor": _to_number(bm.group(2)),
        })

    for pat_str, tipo, desc in [
        (r"Juros de Mora\s+\w+/\d+\s+([\d.,]+)", "juros_mora", "Juros de Mora"),
        (r"Multa por Atraso Pgto\s+\w+/\d+\s+([\d.,]+)", "multa", "Multa por Atraso"),
        (r"Atualiza[cç][aã]o Monet[aá]ria\s+\w+/\d+\s+([\d.,]+)", "atualizacao_monetaria", "Atualizacao Monetaria"),
    ]:
        for rm in re.finditer(pat_str, full_text):
            items.append({"tipo": tipo, "descricao": desc, "valor": _to_number(rm.group(1))})

    data["itens"] = items

    # Total fatura (Total Distribuidora)
    m = re.search(r"Total Distribuidora\s+([\d.,]+)", full_text)
    if m:
        data["total_fatura"] = _to_number(m.group(1))

    # Medidor
    m = re.search(
        r"(\d{8,10})\s+Energia Ativa-kWh [úu]nico\s+(\d+)\s+(\d+)\s+([\d.,]+)\s+(\d+)",
        full_text
    )
    if m:
        data["medidores"] = [{
            "numero": m.group(1),
            "grandeza": "Energia Ativa",
            "leitura_anterior": _to_number(m.group(2)),
            "leitura_atual": _to_number(m.group(3)),
            "constante": _to_number(m.group(4)),
            "consumo": _to_number(m.group(5)),
        }]
        data["uc"] = m.group(1)

    # Rodape: NF Série 0 conta total vencimento
    m = re.search(r"(\d{8,10})\s+S[eé]rie\s+0\s+(\d{10,13})\s+([\d.,]+)\s+(\d{2}/\d{2}/\d{4})", full_text)
    if m:
        if not data.get("nota_fiscal"):
            data["nota_fiscal"] = m.group(1)
        data["conta_contrato"] = m.group(2)
        if not data.get("total_a_pagar"):
            data["total_a_pagar"] = _to_number(m.group(3))
        if not data.get("data_vencimento"):
            data["data_vencimento"] = _to_date(m.group(4))

    return data


# ---------------------------------------------------------------------------
# Funcao principal
# ---------------------------------------------------------------------------

def extract_fatura(pdf_path):
    lines = _extract_horizontal_lines(pdf_path)
    full_text = "\n".join(lines)

    data = {
        "_source_file": str(pdf_path),
        "_extracted_at": datetime.datetime.now().isoformat(),
        "distribuidora": "CPFL Piratininga",
        "modalidade": "Convencional",
    }

    is_danf3e = "DANF3E" in full_text or "COMPANHIA PIRATININGA" in full_text.upper()

    if is_danf3e:
        data["_formato"] = "DANF3E"
        extra = _extract_formato2(lines, full_text)
    else:
        data["_formato"] = "NotaFiscal"
        extra = _extract_formato1(lines, full_text)

    data.update(extra)

    # Bandeira e dias por patamar
    dias_patamar = _extract_bandeira_dias(full_text)
    data["dias_por_patamar"] = dias_patamar
    if not data.get("bandeira_vigente"):
        data["bandeira_vigente"] = _detect_bandeira_vigente(full_text, dias_patamar)

    # GD
    m = re.search(r"Energia\s+injetada\s+no\s+m[eê]s\s+([\d.,]+)\s+kWh", full_text, re.IGNORECASE)
    if m:
        data["gd_injetada_mes"] = _to_number(m.group(1))
    if re.search(r"Unidade\s+Microger[aá]", full_text, re.IGNORECASE):
        data["mmgd_tipo"] = "Microgeracao"
    elif re.search(r"Unidade\s+Miniger[aá]", full_text, re.IGNORECASE):
        data["mmgd_tipo"] = "Minigeracao"
    data["tem_gd"] = bool(data.get("gd_injetada_mes") or data.get("mmgd_tipo"))

    # Consumo faturado / bruto
    for it in data.get("itens", []):
        if it.get("tipo") == "consumo_tusd":
            data["consumo_faturado"] = it.get("quantidade")
            break
    for med in data.get("medidores", []):
        if "Ativa" in med.get("grandeza", ""):
            data["consumo_bruto"] = med.get("consumo")
            break

    # Historico
    historico = []
    for line in lines:
        m = re.match(r"^([A-Z]{3})\s+(\d{2})\s+[l|]+\s+(\d+)\s+(\d+)$", line.strip())
        if m:
            historico.append({
                "mes": m.group(1), "ano": "20" + m.group(2),
                "consumo_kwh": int(m.group(3)), "dias": int(m.group(4)),
            })
    data["historico_consumo"] = historico

    # Impedimento / tipo de leitura (observacoes importantes)
    leitura_aviso = None
    for pat, descricao in [
        (r"impedimento\s+de\s+leitura", "Impedimento de leitura"),
        (r"leitura\s+impedida",          "Leitura impedida"),
        (r"leitura\s+estimada",          "Leitura estimada"),
        (r"leitura\s+pela\s+m[eé]dia",   "Leitura pela média"),
    ]:
        if re.search(pat, full_text, re.IGNORECASE):
            leitura_aviso = descricao
            break
    data["leitura_estimada"] = leitura_aviso is not None
    data["leitura_aviso"] = leitura_aviso

    return data


def summarize(data):
    print(f"\n=== {data.get('_source_file')} ({data.get('_formato', '?')}) ===")
    print(f"Cliente: {data.get('cliente_nome')} | CPF: {data.get('cliente_cpf')}")
    print(f"UC/Medidor: {data.get('uc')} | Cod.Cliente: {data.get('codigo_cliente')}")
    print(f"Subgrupo: {data.get('subgrupo')} | Modalidade: {data.get('modalidade')}")
    print(f"Periodo: {data.get('leitura_anterior')} a {data.get('leitura_atual')} ({data.get('dias_ciclo')} dias)")
    print(f"Mes ref: {data.get('mes_ref')} | Vencimento: {data.get('data_vencimento')}")
    print(f"NF: {data.get('nota_fiscal')} | Emissao: {data.get('data_emissao')}")
    print(f"Bandeira: {data.get('bandeira_vigente')} | Dias: {data.get('dias_por_patamar')}")
    print(f"Consumo bruto: {data.get('consumo_bruto')} | faturado: {data.get('consumo_faturado')}")
    print(f"Total fatura: {data.get('total_fatura')} | A pagar: {data.get('total_a_pagar')}")
    print(f"Itens ({len(data.get('itens', []))}):")
    for it in data.get("itens", []):
        print(f"   {it.get('tipo'):25s} | qtd={it.get('quantidade','-')} | val=R${it.get('valor','-')}")
    trib = {k: v for k, v in data.get('tributos', {}).items()}
    print(f"Tributos: {trib}")


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Uso: python extractor.py <fatura.pdf>")
        sys.exit(1)
    d = extract_fatura(sys.argv[1])
    summarize(d)

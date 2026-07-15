# -*- coding: utf-8 -*-
"""
audit.py - Logica de triagem para faturas Neoenergia PE
Faixas: OK / INVESTIGAR / DIVERGENCIA
"""

# Tolerancias
TOL_ITEM   = 0.10   # R$ 0,10 para math de item (qtd x preco = valor)
TOL_ICMS   = 0.10   # R$ 0,10 para math de ICMS/PIS/COFINS
TOL_LEIT   = 2      # kWh de tolerancia na leitura do medidor
TOL_SOMA_OK   = 0.10  # fracao: soma itens vs total dentro de 10% => OK
TOL_SOMA_INV  = 0.30  # fracao: acima de 30% => DIVERGENCIA (ajuste grande)

# Campos obrigatorios minimos
CAMPOS_CRIT = [
    "ref_mes_ano", "vencimento", "conta_contrato",
    "consumo_kwh_tusd_qtd", "valor_tusd", "valor_te",
    "preco_tusd", "preco_te", "total_fatura",
]
CAMPOS_SEC = [
    "cod_instalacao", "data_emissao", "nr_medidor",
    "leitura_anterior", "leitura_atual", "cosip", "icms_aliq",
]


def _chk(val, tol):
    """True se val <= tol, None se val indefinido."""
    if val is None:
        return None
    return abs(val) <= tol


def auditar(r):
    """
    Recebe dict de parse_fatura e devolve:
      triagem   : "OK" | "INVESTIGAR" | "DIVERGENCIA"
      motivos   : list[str] com descricao dos problemas
      metricas  : dict com valores calculados uteis para o Excel
    """
    motivos  = []
    metricas = {}
    flags_div  = []   # forcam DIVERGENCIA
    flags_inv  = []   # forcam INVESTIGAR (se nao houver DIVERGENCIA)

    # ── 1. Campos criticos ausentes ──────────────────────────────────────
    faltam_crit = [c for c in CAMPOS_CRIT if r.get(c) is None]
    faltam_sec  = [c for c in CAMPOS_SEC  if r.get(c) is None]
    if faltam_crit:
        flags_div.append(f"campos criticos ausentes: {', '.join(faltam_crit)}")
    elif faltam_sec:
        flags_inv.append(f"campos secundarios ausentes: {', '.join(faltam_sec)}")

    # ── 2. Math TUSD: qtd x preco = valor ───────────────────────────────
    qtd   = r.get("consumo_kwh_tusd_qtd")
    ptsd  = r.get("preco_tusd")
    vtusd = r.get("valor_tusd")
    if qtd and ptsd and vtusd:
        calc_tusd = round(qtd * ptsd, 2)
        dif_tusd  = abs(calc_tusd - vtusd)
        metricas["calc_TUSD"]   = calc_tusd
        metricas["dif_TUSD_R$"] = round(dif_tusd, 2)
        if dif_tusd > TOL_ITEM:
            flags_div.append(
                f"TUSD math: {qtd}kWh x {ptsd} = {calc_tusd:.2f} != {vtusd} "
                f"(dif={dif_tusd:.2f})"
            )
    else:
        metricas["dif_TUSD_R$"] = None
        if not faltam_crit:
            flags_inv.append("TUSD: qtd/preco/valor ausente")

    # ── 3. Math TE: qtd x preco = valor ─────────────────────────────────
    pte  = r.get("preco_te")
    vte  = r.get("valor_te")
    qtde = r.get("consumo_kwh_te_qtd") or qtd   # geralmente igual ao TUSD
    if qtde and pte and vte:
        calc_te = round(qtde * pte, 2)
        dif_te  = abs(calc_te - vte)
        metricas["calc_TE"]   = calc_te
        metricas["dif_TE_R$"] = round(dif_te, 2)
        if dif_te > TOL_ITEM:
            flags_div.append(
                f"TE math: {qtde}kWh x {pte} = {calc_te:.2f} != {vte} "
                f"(dif={dif_te:.2f})"
            )
    else:
        metricas["dif_TE_R$"] = None

    # ── 4. Leitura do medidor ────────────────────────────────────────────
    lant = r.get("leitura_anterior")
    latu = r.get("leitura_atual")
    cte  = r.get("constante_medidor") or 1.0
    if lant is not None and latu is not None and qtd is not None:
        calc_leit = round((latu - lant) * cte, 1)
        dif_leit  = abs(calc_leit - qtd)
        metricas["calc_leit_kWh"] = calc_leit
        metricas["dif_leit_kWh"]  = round(dif_leit, 1)
        if dif_leit > TOL_LEIT:
            # Verifica mínimo (Custo de Disponibilidade) por tipo de fornecimento
            _tipo = (r.get("tipo_fornecimento") or "").lower()
            if "trifasico" in _tipo or "trifásico" in _tipo:
                _minimo = 100
            elif "bifasico" in _tipo or "bifásico" in _tipo:
                _minimo = 30 if ("dois" in _tipo or "2 " in _tipo) else 50
            elif "monofasico" in _tipo or "monofásico" in _tipo:
                _minimo = 30
            else:
                _minimo = None

            if _minimo is not None and calc_leit < _minimo and round(qtd, 1) == _minimo:
                # Faturamento pelo mínimo correto — não é divergência
                metricas["custo_disponibilidade"] = _minimo
            else:
                flags_inv.append(
                    f"leitura: ({latu}-{lant})x{cte}={calc_leit} "
                    f"!= consumo={qtd} (dif={dif_leit:.1f}kWh)"
                )
    else:
        metricas["dif_leit_kWh"] = None

    # ── 5. Math ICMS ─────────────────────────────────────────────────────
    icms_b = r.get("icms_base")
    icms_a = r.get("icms_aliq")
    icms_v = r.get("icms_valor")
    if icms_b and icms_a and icms_v:
        calc_icms = round(icms_b * icms_a / 100, 2)
        dif_icms  = abs(calc_icms - icms_v)
        metricas["calc_ICMS"]   = calc_icms
        metricas["dif_ICMS_R$"] = round(dif_icms, 2)
        if dif_icms > TOL_ICMS:
            flags_inv.append(
                f"ICMS math: {icms_b}x{icms_a}%={calc_icms:.2f} != {icms_v} "
                f"(dif={dif_icms:.2f})"
            )
    else:
        metricas["dif_ICMS_R$"] = None

    # ── 5b. SCEE — auditoria da compensação ─────────────────────────────
    if r.get("is_scee"):
        scee_kwh  = r.get("scee_kwh_compensados") or 0
        comp_c    = r.get("valor_imp_som_dim_c") or 0  # negativo
        preco_tot = (r.get("preco_tusd") or 0) + (r.get("preco_te") or 0)
        metricas["scee_kwh_compensados"] = scee_kwh
        metricas["is_scee"] = True
        if scee_kwh and preco_tot:
            comp_aud = round(scee_kwh * preco_tot, 2)
            comp_cob = round(abs(comp_c), 2)
            dif_scee = abs(comp_aud - comp_cob)
            metricas["comp_scee_auditado_R$"] = comp_aud
            metricas["comp_scee_cobrado_R$"]  = comp_cob
            metricas["dif_scee_R$"]           = round(dif_scee, 2)
            status_scee = "OK" if dif_scee <= TOL_ITEM else "INVESTIGAR"
            if status_scee == "INVESTIGAR":
                flags_inv.append(
                    f"SCEE: {scee_kwh}kWh × R${preco_tot:.6f}/kWh "
                    f"= R${comp_aud:.2f} auditado vs R${comp_cob:.2f} cobrado "
                    f"(dif=R${dif_scee:.2f})"
                )
        else:
            flags_inv.append("SCEE detectado mas dados insuficientes para auditar compensacao")

    # ── 6. Soma dos itens vs total da fatura ─────────────────────────────
    total = r.get("total_fatura")
    soma  = sum(
        v for v in [
            r.get("valor_tusd"),
            r.get("valor_te"),
            r.get("valor_bandeira") or 0,
            r.get("cosip") or 0,
            r.get("icms_cde") or 0,
            r.get("valor_parcelamento") or 0,
            r.get("valor_ipca") or 0,
            r.get("valor_imp_som_dim_c") or 0,   # SCEE crédito (negativo)
            r.get("valor_imp_som_dim_s") or 0,   # SCEE ajuste sem imposto
        ]
        if v is not None
    )
    metricas["soma_itens_R$"] = round(soma, 2)
    if total and soma:
        dif_total = soma - total          # positivo = soma > total (credito nao extraido)
        dif_pct   = abs(dif_total) / total if total else 0
        metricas["dif_total_R$"] = round(dif_total, 2)
        metricas["dif_total_%"]  = round(dif_pct * 100, 1)

        if dif_pct > TOL_SOMA_INV:
            flags_div.append(
                f"soma itens R${soma:.2f} != total R${total:.2f} "
                f"(dif={dif_total:+.2f}, {dif_pct*100:.0f}%) "
                f"-- possivelmente ajuste/credito nao extraido"
            )
        elif dif_pct > TOL_SOMA_OK:
            flags_inv.append(
                f"soma itens R${soma:.2f} != total R${total:.2f} "
                f"(dif={dif_total:+.2f}, {dif_pct*100:.0f}%) "
                f"-- possivel multa/juros nao extraidos"
            )
    else:
        metricas["dif_total_R$"] = None
        metricas["dif_total_%"]  = None

    # ── 7. Erros de extracao ─────────────────────────────────────────────
    if r.get("erro"):
        flags_div.append(f"erro de extracao: {r['erro']}")

    # ── Triagem final ─────────────────────────────────────────────────────
    if flags_div:
        triagem = "DIVERGENCIA"
        motivos = flags_div + flags_inv
    elif flags_inv:
        triagem = "INVESTIGAR"
        motivos = flags_inv
    else:
        triagem = "OK"
        motivos = []

    return triagem, motivos, metricas

# -*- coding: utf-8 -*-
"""
cpfl_paulista/audit.py
Regras de auditoria para CPFL Paulista — Tarifa Verde A4
"""
import json
from datetime import datetime
from pathlib import Path

_DATA_DIR = Path(__file__).parent / "data"
_REHS = None
_DATA_NOVA_FORMULA_ICMS = datetime(2023, 3, 1)  # mantido apenas como referencia


def _detectar_formula_stf(r):
    """
    Detecta se a fatura aplica a formula STF RE 574.706 (ICMS excluido da base PIS/COFINS).
    Criterio: base ICMS != base PIS/COFINS na fatura -> formula nova (STF).
              base ICMS == base PIS/COFINS -> formula antiga.
    """
    icms_b = r.get("icms_base")
    pis_b  = r.get("pis_base")
    if icms_b is None or pis_b is None:
        return False
    return abs(icms_b - pis_b) > 0.01


def _carregar_rehs():
    global _REHS
    if _REHS is None:
        with open(_DATA_DIR / "rehs.json", encoding="utf-8") as f:
            _REHS = json.load(f)
    return _REHS


def _reh_para_periodo(ref_mes_ano):
    """Retorna a tarifa REH vigente para o ref_mes_ano (ex: 'JAN/2026')."""
    meses = {"JAN":1,"FEV":2,"MAR":3,"ABR":4,"MAI":5,"JUN":6,
             "JUL":7,"AGO":8,"SET":9,"OUT":10,"NOV":11,"DEZ":12}
    partes = ref_mes_ano.split("/")
    if len(partes) != 2:
        return None
    mes = meses.get(partes[0].upper())
    ano = int(partes[1])
    if not mes:
        return None
    ref_dt = datetime(ano, mes, 1)

    rehs = _carregar_rehs()
    for t in rehs["tarifas"]:
        ini = datetime.strptime(t["vigencia_inicio"], "%Y-%m-%d")
        fim = datetime.strptime(t["vigencia_fim"], "%Y-%m-%d")
        if ini <= ref_dt <= fim:
            return t
    return None


def _parse_ref_cpfl(ref_mes_ano):
    """Converte 'JAN/2026' para datetime(2026, 1, 1) ou None."""
    meses = {"JAN":1,"FEV":2,"MAR":3,"ABR":4,"MAI":5,"JUN":6,
             "JUL":7,"AGO":8,"SET":9,"OUT":10,"NOV":11,"DEZ":12}
    if not ref_mes_ano:
        return None
    partes = ref_mes_ano.split("/")
    if len(partes) != 2:
        return None
    mes = meses.get(partes[0].upper())
    try:
        ano = int(partes[1])
    except ValueError:
        return None
    if not mes:
        return None
    return datetime(ano, mes, 1)


def _base_icms_esperada_cpfl(net, icms_aliq_pct, pis_aliq_pct, cofins_aliq_pct, ref_mes_ano):
    """
    Calcula base ICMS esperada — CPFL Paulista (MT Verde A4).
    net: soma dos componentes tributáveis (TUSD + TE + demanda utilizada + bandeira, sem COSIP).
    - A partir de 03/2023 (STF RE 574.706):
        base = net / ((1 - ICMS) * (1 - (PIS + COFINS)))
    - Antes de 03/2023:
        base = net / (1 - (ICMS + PIS + COFINS))
    Retorna (base_R$, nova_formula: bool) ou (None, None).
    """
    if net <= 0 or not icms_aliq_pct or not pis_aliq_pct or not cofins_aliq_pct:
        return None, None
    icms   = icms_aliq_pct   / 100
    pis    = pis_aliq_pct    / 100
    cofins = cofins_aliq_pct / 100
    ref_dt = _parse_ref_cpfl(ref_mes_ano)
    nova_formula = (ref_dt is not None and ref_dt >= _DATA_NOVA_FORMULA_ICMS)
    denom = (1 - icms) * (1 - (pis + cofins)) if nova_formula else 1 - (icms + pis + cofins)
    if denom <= 0:
        return None, None
    return round(net / denom, 2), nova_formula


def _tol(valor_cobrado, valor_esperado, tolerancia=0.0005):
    """True se diferença relativa > tolerância."""
    if valor_esperado is None or valor_esperado == 0:
        return False
    return abs(valor_cobrado - valor_esperado) / abs(valor_esperado) > tolerancia


def auditar(r):
    """
    r: dict retornado por extractor.parse_fatura()
    Retorna: (triagem, motivos, metricas)
      triagem  : "OK" | "INVESTIGAR" | "DIVERGENCIA"
      motivos  : list[str]
      metricas : dict
    """
    alertas = []
    metricas = {}

    reh = _reh_para_periodo(r.get("ref_mes_ano", ""))

    # ── 1. REH — Tarifas sem tributação ──────────────────────────────────────
    if reh:
        p = reh["postos"]
        checks = [
            ("TUSD Ponta",     r.get("tusd_ponta_sem"),   p["Ponta"].get("TUSD_kwh")),
            ("TUSD Fora Ponta",r.get("tusd_fp_sem"),      p["ForaPonta"].get("TUSD_kwh")),
            ("TE Ponta",       r.get("te_ponta_sem"),      p["Ponta"].get("TE_kwh")),
            ("TE Fora Ponta",  r.get("te_fp_sem"),         p["ForaPonta"].get("TE_kwh")),
            ("Demanda TUSD",   r.get("tusd_demanda_sem"),  p["Ponta"].get("TUSD_kw")),
        ]
        for nome, cobrado, esperado in checks:
            if cobrado is None or esperado is None:
                continue
            dif = round(cobrado - esperado, 8)
            if _tol(cobrado, esperado):
                alertas.append({
                    "cat": "Tarifa",
                    "descricao": f"{nome}: cobrado {cobrado:.8f} vs REH {esperado:.8f} (dif {dif:+.8f})",
                    "nivel": "DIVERGENCIA",
                    "diferenca": dif,
                })
        metricas["reh_aplicado"] = reh.get("reh")
    else:
        alertas.append({
            "cat": "REH",
            "descricao": f"Nenhuma REH cadastrada para o período {r.get('ref_mes_ano')}",
            "nivel": "INVESTIGAR",
        })

    # ── 2. Período de leitura — RN 1000/2021 Art. 261 (MT) ───────────────────
    # Normal: mês civil (28-31 dias); excepcional: 15-47 dias
    nr_dias = r.get("nr_dias")
    if nr_dias is not None:
        metricas["nr_dias"] = nr_dias
        if nr_dias < 15 or nr_dias > 47:
            alertas.append({
                "cat": "Período",
                "descricao": (
                    f"Período de faturamento: {nr_dias} dias fora do limite legal "
                    f"(RN 1000/2021 Art. 261: mínimo 15, máximo 47 dias)"
                ),
                "nivel": "DIVERGENCIA",
            })
        elif nr_dias < 28 or nr_dias > 31:
            alertas.append({
                "cat": "Período",
                "descricao": (
                    f"Período de faturamento MT: {nr_dias} dias fora do mês civil "
                    f"(RN 1000/2021 Art. 261: normal = mês civil; "
                    f"15-47 apenas em casos excepcionais)"
                ),
                "nivel": "INVESTIGAR",
            })

    # ── 3. Demanda ultrapassagem ──────────────────────────────────────────────
    dem_cont   = r.get("demanda_contratada_kw") or 0
    dem_med    = r.get("demanda_medida_kw") or 0
    dem_ultrap = r.get("demanda_ultrap_kw")
    if dem_ultrap and dem_cont > 0:
        pct_ultrap = dem_ultrap / dem_cont * 100
        metricas["demanda_ultrapassagem_pct"] = round(pct_ultrap, 1)
        if pct_ultrap > 5:
            alertas.append({
                "cat": "Demanda",
                "descricao": (f"Ultrapassagem de demanda: {dem_ultrap:.1f} kW "
                              f"({pct_ultrap:.1f}% acima do contratado {dem_cont:.0f} kW) — "
                              f"verifique necessidade de adequação contratual"),
                "nivel": "INVESTIGAR",
            })

    # ── 4. ICMS — alíquota esperada 18% ──────────────────────────────────────
    icms_aliq = r.get("icms_aliq")
    icms_b    = r.get("icms_base")
    if icms_aliq is not None and abs(icms_aliq - 18.0) > 0.01:
        alertas.append({
            "cat": "ICMS",
            "descricao": f"Alíquota ICMS cobrada {icms_aliq}% — esperado 18% (SP)",
            "nivel": "DIVERGENCIA",
        })

    # ── 4b. Base de cálculo do ICMS — composição correta (MT Verde A4) ───────
    # Base ICMS = soma direta dos valores COM tributos:
    #   consumo TUSD ponta+FP + consumo TE ponta+FP (bruto, antes dos créditos)
    #   + reativo exc ponta+FP
    #   + valor_demanda (kW medido × tusd_com, conforme fatura)
    #   + valor_demanda_ultrap
    #   + bandeira
    #   − créditos SCEE TE ponta+FP (reduzem a base; TUSD injetado tem base=0)
    # COSIP e USDG NÃO entram.
    pis_aliq    = r.get("pis_aliq")
    cofins_aliq = r.get("cofins_aliq")
    if icms_b and icms_aliq and pis_aliq and cofins_aliq:
        # Pré-calcula denominador para 4c (tarifa COM check)
        nova_f_cpfl  = _detectar_formula_stf(r)
        icms_c   = icms_aliq   / 100
        pis_c    = pis_aliq    / 100
        cofins_c = cofins_aliq / 100
        denom_cpfl = ((1 - icms_c) * (1 - (pis_c + cofins_c)) if nova_f_cpfl
                      else 1 - (icms_c + pis_c + cofins_c))
        metricas["formula_icms"] = "STF/2021" if nova_f_cpfl else "anterior_03-2023"

        # Soma direta dos valores COM — sem aplicar fórmula de grossup
        # Créditos SCEE TE deduzem a base; créditos TUSD têm base ICMS = 0
        base_esp = round(
            (r.get("valor_tusd_ponta")       or 0)
            + (r.get("valor_tusd_fp")        or 0)
            + (r.get("valor_te_ponta")       or 0)
            + (r.get("valor_te_fp")          or 0)
            + (r.get("valor_demanda")        or 0)
            + (r.get("valor_demanda_ultrap") or 0)
            + (r.get("valor_reativo_exc_ponta") or 0)
            + (r.get("valor_reativo_exc_fp")    or 0)
            + (r.get("valor_bandeira")       or 0)
            - (r.get("valor_inj_fp_te")      or 0)
            - (r.get("valor_inj_ponta_te")   or 0),
            2
        )
        if base_esp > 0:
            dif_base     = abs(base_esp - icms_b)
            dif_base_pct = dif_base / base_esp if base_esp else 0
            metricas["base_icms_esperada"] = base_esp
            metricas["dif_base_icms_R$"]   = round(dif_base, 2)
            if dif_base_pct > 0.001:
                alertas.append({
                    "cat": "ICMS",
                    "descricao": (
                        f"Base ICMS: fatura={icms_b:.2f} vs esperado={base_esp:.2f} "
                        f"(dif={icms_b - base_esp:+.2f}) "
                        f"— composição incorreta: esperado consumo+reativo+demanda+ultrap+bandeira"
                        f"−SCEE_TE, sem COSIP nem USDG"
                    ),
                    "nivel": "DIVERGENCIA",
                })

    # ── 4c. Tarifa COM tributos — verificar cálculo a partir da tarifa SEM ────
    # Formula nova (STF RE 574.706, >= 03/2023): tarifa_com = tarifa_sem / ((1-ICMS)*(1-PIS-COFINS))
    # Formula antiga (< 03/2023):                tarifa_com = tarifa_sem / (1-ICMS-PIS-COFINS)
    if icms_aliq and pis_aliq and cofins_aliq:
        if "denom_cpfl" not in locals():
            nova_f_cpfl  = _detectar_formula_stf(r)
            icms_c   = icms_aliq   / 100
            pis_c    = pis_aliq    / 100
            cofins_c = cofins_aliq / 100
            denom_cpfl = ((1 - icms_c) * (1 - (pis_c + cofins_c)) if nova_f_cpfl
                          else 1 - (icms_c + pis_c + cofins_c))
        if denom_cpfl > 0:
            for nome_t, key_sem_t, key_com_t in [
                ("TUSD Ponta",    "tusd_ponta_sem",    "tusd_ponta_com"),
                ("TUSD FP",       "tusd_fp_sem",       "tusd_fp_com"),
                ("TE Ponta",      "te_ponta_sem",      "te_ponta_com"),
                ("TE FP",         "te_fp_sem",         "te_fp_com"),
                ("Demanda TUSD",  "tusd_demanda_sem",  "tusd_demanda_com"),
            ]:
                tar_sem_t = r.get(key_sem_t)
                tar_com_t = r.get(key_com_t)
                if tar_sem_t and tar_com_t:
                    preco_exp_t = round(tar_sem_t / denom_cpfl, 8)
                    dif_pct_t   = abs(tar_com_t - preco_exp_t) / preco_exp_t if preco_exp_t else 0
                    metricas[f"preco_com_{nome_t.replace(' ','_')}_esp"] = preco_exp_t
                    if dif_pct_t > 0.0005:
                        alertas.append({
                            "cat": "ICMS",
                            "descricao": (
                                f"Tarifa COM {nome_t}: fatura={tar_com_t:.8f} "
                                f"vs esperado={preco_exp_t:.8f} "
                                f"(sem={tar_sem_t:.8f}, dif={tar_com_t - preco_exp_t:+.8f}, "
                                f"fórmula {'nova STF' if nova_f_cpfl else 'antiga'})"
                            ),
                            "nivel": "DIVERGENCIA",
                        })

    # ── 5. Total a pagar vs total fatura ──────────────────────────────────────
    total_fat   = r.get("total_fatura") or 0
    total_pagar = r.get("total_a_pagar") or 0
    dif_total   = round(total_pagar - total_fat, 2)
    metricas["dif_total_R$"] = dif_total
    if abs(dif_total) > 0.05:
        alertas.append({
            "cat": "Total",
            "descricao": f"Total a pagar R$ {total_pagar:,.2f} ≠ total fatura R$ {total_fat:,.2f} (dif R$ {dif_total:+,.2f})",
            "nivel": "DIVERGENCIA",
        })

    # ── 6. Leituras — consumo medido vs faturado ─────────────────────────────
    taxa_perda  = r.get("taxa_perda", False)
    fator_perda = 1.025 if taxa_perda else 1.0
    TOL_ABS     = 2.0  # tolerância absoluta (arredondamento distribuidora)

    leit_checks = [
        ("kWh Ponta",        r.get("med_kwh_ponta_cons"),   r.get("consumo_ponta_kwh")),
        ("kWh Fora Ponta",   r.get("med_kwh_fp_cons"),      r.get("consumo_fp_kwh")),
        ("kW Demanda",       r.get("med_kw_fp_cons"),        r.get("demanda_medida_kw")),
        ("kVarh Ponta",      r.get("med_kvarh_ponta_cons"), r.get("consumo_reativo_exc_ponta_kwh")),
        ("kVarh Fora Ponta", r.get("med_kvarh_fp_cons"),    r.get("consumo_reativo_exc_fp_kwh")),
    ]
    for nome_leit, cons_med, fat in leit_checks:
        if cons_med is None or fat is None:
            continue
        esperado = round(cons_med * fator_perda, 3)
        dif = round(esperado - fat, 3)
        if abs(dif) > TOL_ABS:
            alertas.append({
                "cat": "Leitura",
                "descricao": (
                    f"{nome_leit}: medido={cons_med:,.3f}"
                    f"{'×1,025' if taxa_perda else ''} = {esperado:,.3f}"
                    f" vs faturado {fat:,.4f} (dif {dif:+,.3f})"
                ),
                "nivel": "DIVERGENCIA",
            })
    metricas["taxa_perda"] = taxa_perda

    # ── 8. SCEE — verificação de créditos de compensação ─────────────────────
    inj_fp    = r.get("injetada_fp_kwh") or 0
    inj_ponta = r.get("injetada_ponta_kwh") or 0

    if inj_fp > 0 or inj_ponta > 0:
        pis_c    = (pis_aliq    or 0) / 100
        cofins_c = (cofins_aliq or 0) / 100
        icms_c   = (icms_aliq   or 0) / 100
        denom_tusd = 1 - pis_c - cofins_c
        denom_te   = (1 - icms_c) * (1 - pis_c - cofins_c)

        dif_scee_total = 0
        checks_scee = [
            ("TUSD FP",    inj_fp,    r.get("tusd_fp_sem"),    r.get("valor_inj_fp_tusd"),    denom_tusd),
            ("TUSD Ponta", inj_ponta, r.get("tusd_ponta_sem"), r.get("valor_inj_ponta_tusd"), denom_tusd),
            ("TE FP",      inj_fp,    r.get("te_fp_sem"),      r.get("valor_inj_fp_te"),      denom_te),
            ("TE Ponta",   inj_ponta, r.get("te_ponta_sem"),   r.get("valor_inj_ponta_te"),   denom_te),
        ]
        for nome, kwh, sem, val_cobrado, denom in checks_scee:
            if kwh <= 0 or sem is None or not val_cobrado or denom <= 0:
                continue
            val_esp = round(kwh * sem / denom, 2)
            val_cob = round(val_cobrado, 2)
            dif     = round(val_esp - val_cob, 2)   # positivo = cliente recebeu menos crédito
            metricas[f"scee_{nome.replace(' ', '_').lower()}_esp"] = val_esp
            if abs(dif) > 0.05:
                dif_scee_total += dif
                alertas.append({
                    "cat": "SCEE",
                    "descricao": (
                        f"Crédito SCEE {nome}: fatura=R${val_cob:.2f} "
                        f"vs esperado=R${val_esp:.2f} "
                        f"({kwh:.4f} kWh × {sem:.8f} / {denom:.6f}) "
                        f"dif=R${dif:+.2f}"
                    ),
                    "nivel": "DIVERGENCIA",
                })
        if dif_scee_total != 0:
            metricas["dif_scee_R$"] = round(dif_scee_total, 2)

        # ── USDG — ICMS deve ser zero; não deve entrar na base de cálculo ────
        # A base_esp na seção 4b não inclui USDG. Se a distribuidora incluir
        # USDG na base ICMS, a seção 4b flagrará DIVERGÊNCIA automaticamente.
        valor_usdg        = r.get("valor_usdg") or 0
        valor_usdg_ultrap = r.get("valor_usdg_ultrap") or 0
        if valor_usdg > 0 or valor_usdg_ultrap > 0:
            metricas["valor_usdg"]        = valor_usdg
            metricas["valor_usdg_ultrap"] = valor_usdg_ultrap
            # Demanda de geração contratada — verificar se informada na fatura
            dem_ger_cont = r.get("demanda_geracao_contratada_kw")
            if dem_ger_cont is None:
                alertas.append({
                    "cat": "USDG",
                    "descricao": (
                        "Demanda de geração contratada não localizada na fatura "
                        f"(USDG={valor_usdg:.2f} | Ultrap={valor_usdg_ultrap:.2f}) "
                        "— verificar contrato e conferir manualmente"
                    ),
                    "nivel": "INVESTIGAR",
                })

    # ── Triagem final ─────────────────────────────────────────────────────────
    niveis = [a.get("nivel", "OK") for a in alertas]
    if "DIVERGENCIA" in niveis:
        triagem = "DIVERGENCIA"
    elif "INVESTIGAR" in niveis:
        triagem = "INVESTIGAR"
    else:
        triagem = "OK"

    motivos = [a["descricao"] for a in alertas if a.get("nivel") != "OK"]

    return triagem, motivos, metricas

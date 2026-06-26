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

    # ── 2. Período de leitura (15–45 dias) ────────────────────────────────────
    nr_dias = r.get("nr_dias")
    if nr_dias is not None:
        if nr_dias < 15 or nr_dias > 45:
            alertas.append({
                "cat": "Período",
                "descricao": f"Ciclo de {nr_dias} dias fora do limite regulatório (15–45 dias) — REN 1.000/2021",
                "nivel": "INVESTIGAR",
            })

    # ── 3. Demanda ultrapassagem ──────────────────────────────────────────────
    dem_cont = r.get("demanda_contratada_kw") or 0
    dem_med  = r.get("demanda_medida_kw") or 0
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
    if icms_aliq is not None and abs(icms_aliq - 18.0) > 0.01:
        alertas.append({
            "cat": "ICMS",
            "descricao": f"Alíquota ICMS cobrada {icms_aliq}% — esperado 18% (SP)",
            "nivel": "DIVERGENCIA",
        })

    # ── 5. Total a pagar vs total fatura ──────────────────────────────────────
    total_fat  = r.get("total_fatura") or 0
    total_pagar = r.get("total_a_pagar") or 0
    dif_total  = round(total_pagar - total_fat, 2)
    metricas["dif_total_R$"] = dif_total
    if abs(dif_total) > 0.05:
        alertas.append({
            "cat": "Total",
            "descricao": f"Total a pagar R$ {total_pagar:,.2f} ≠ total fatura R$ {total_fat:,.2f} (dif R$ {dif_total:+,.2f})",
            "nivel": "DIVERGENCIA",
        })

    # ── 6. Leituras — consumo medido vs faturado ─────────────────────────────
    taxa_perda = r.get("taxa_perda", False)
    fator_perda = 1.025 if taxa_perda else 1.0
    TOL_ABS = 2.0  # tolerância absoluta (arredondamento distribuidora)

    leit_checks = [
        ("kWh Ponta",    r.get("med_kwh_ponta_cons"),   r.get("consumo_ponta_kwh")),
        ("kWh Fora Ponta", r.get("med_kwh_fp_cons"),    r.get("consumo_fp_kwh")),
        ("kW Demanda",   r.get("med_kw_fp_cons"),        r.get("demanda_medida_kw")),
        ("kVarh Ponta",  r.get("med_kvarh_ponta_cons"),  None),   # sem campo faturado direto ainda
        ("kVarh Fora Ponta", r.get("med_kvarh_fp_cons"), None),
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

    # ── 8. GD — compensação TE FP ────────────────────────────────────────────
    inj_fp  = r.get("injetada_fp_kwh") or 0
    cons_fp = r.get("consumo_fp_kwh") or 0
    if inj_fp > 0 and inj_fp > cons_fp:
        alertas.append({
            "cat": "GD",
            "descricao": f"Energia injetada FP ({inj_fp:.1f} kWh) > consumo FP ({cons_fp:.1f} kWh) — verificar compensação",
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

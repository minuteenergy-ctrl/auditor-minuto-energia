"""
Motor de auditoria de faturas CPFL Piratininga.
Recalcula tarifas com gross-up, bandeira proporcional, Fio B (Lei 14.300/2022),
Tese do Seculo (RE 574.706/STF) e valida varios itens.
"""
import json
import datetime
from pathlib import Path

DATA_DIR = Path(__file__).parent / "data"


def _load_json(fname):
    with open(DATA_DIR / fname, encoding="utf-8") as f:
        return json.load(f)


def _to_date(v):
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, str):
        for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
            try:
                return datetime.datetime.strptime(v, fmt).date()
            except ValueError:
                pass
    return None


def find_tarifa(dados, rehs=None):
    """Retorna o dict de tarifas sem tributos da REH vigente na data da fatura."""
    if rehs is None:
        rehs = _load_json("rehs.json")
    data_ref = dados.get("leitura_atual") or dados.get("data_emissao")
    if data_ref is None:
        return None, None
    data_ref = _to_date(data_ref)
    subgrupo = (dados.get("subgrupo") or "B1").upper()
    sub_base = subgrupo[:2]

    for tarifa in rehs.get("tarifas", []):
        vi = _to_date(tarifa["vigencia_inicio"])
        vf = _to_date(tarifa["vigencia_fim"])
        subs = [s.upper() for s in tarifa.get("subgrupos", [])]
        if vi and vf and vi <= data_ref <= vf and (sub_base in subs or subgrupo in subs):
            postos = tarifa.get("postos", {})
            posto = postos.get("Unico") or postos.get("unico") or next(iter(postos.values()), None)
            return posto, tarifa.get("reh", "")
    return None, None


def find_bandeira(data_ref, bandeiras=None):
    """Retorna dict {patamar: tarifa_r$/kWh} vigente na data."""
    if bandeiras is None:
        bandeiras = _load_json("bandeiras.json")
    if data_ref is None:
        return {}
    data_ref = _to_date(data_ref)
    for b in bandeiras.get("tarifas", []):
        vi = _to_date(b["vigencia_inicio"])
        vf = _to_date(b["vigencia_fim"])
        if vi and vf and vi <= data_ref <= vf:
            return b.get("patamares", {})
    return {}


def grossup(tarifa_sem, icms, pis, cofins):
    """Calcula tarifa com gross-up de tributos."""
    denom = (1 - icms) * (1 - pis - cofins)
    if denom <= 0:
        return tarifa_sem
    return tarifa_sem / denom


def calcular_fio_b(data_ref, data_adesao_mmgd, fio_b_data=None):
    """
    Retorna fator Fio B (0.0 a 1.0) conforme Lei 14.300/2022.
    - Pre-MMGD (adesao <= 06/01/2023): isento (fator=0) ate 2045
    - Pos-MMGD: fator gradual por ano calendario
    """
    if fio_b_data is None:
        fio_b_data = _load_json("fioB_factors.json")
    marco = _to_date(fio_b_data.get("marco_lei", "2023-01-06"))
    data_adesao = _to_date(data_adesao_mmgd)
    if data_adesao is None or data_adesao <= marco:
        return 0.0
    data_ref = _to_date(data_ref)
    if data_ref is None:
        return 0.0
    ano = str(data_ref.year)
    fatores = fio_b_data.get("fatores_pos_mmgd", {})
    return fatores.get(ano) or fatores.get("2029+", 1.0)


def summary_alertas(alertas):
    counts = {"OK": 0, "ATENCAO": 0, "INVESTIGAR": 0}
    for a in alertas:
        s = a.get("status", "OK")
        counts[s] = counts.get(s, 0) + 1
    return counts


def auditar_fatura(dados, config):
    """
    Executa todas as verificacoes de auditoria.
    config keys:
      data_adesao_mmgd (str YYYY-MM-DD)
      tem_gd (bool)
      energia_compensada_kwh (float|None)
      usar_cat_como_compensada (bool)
    Retorna dict com chaves: auditado, cobrado, diferenca, alertas
    """
    rehs = _load_json("rehs.json")
    bandeiras = _load_json("bandeiras.json")
    fio_b_data = _load_json("fioB_factors.json")
    params = _load_json("parametros.json")

    tol_rs = params["tolerancias"]["valor_rs"]
    tol_pct = params["tolerancias"]["tarifa_unit_pct"]
    icms_sp = params["icms_sp"]

    alertas = []
    auditado = {}
    cobrado = {}

    data_ref = dados.get("leitura_atual") or dados.get("data_emissao")

    # ----------------------------------------------------------------
    # 1. Tarifa da REH
    # ----------------------------------------------------------------
    posto_reh, reh_id = find_tarifa(dados, rehs)
    auditado["reh_aplicada"] = reh_id or "N/A"

    trib = dados.get("tributos", {})
    pis_alq = (trib.get("pis", {}).get("aliquota_pct") or 0) / 100
    cofins_alq = (trib.get("cofins", {}).get("aliquota_pct") or 0) / 100
    icms_alq = (trib.get("icms", {}).get("aliquota_pct") or icms_sp * 100) / 100

    if posto_reh:
        tusd_sem = posto_reh.get("TUSD", 0)
        te_sem = posto_reh.get("TE", 0)
        tusd_com = grossup(tusd_sem, icms_alq, pis_alq, cofins_alq)
        te_com = grossup(te_sem, icms_alq, pis_alq, cofins_alq)
        auditado["tusd_sem_trib"] = round(tusd_sem, 8)
        auditado["te_sem_trib"] = round(te_sem, 8)
        auditado["tusd_com_trib"] = round(tusd_com, 8)
        auditado["te_com_trib"] = round(te_com, 8)
    else:
        tusd_com = te_com = tusd_sem = te_sem = 0
        alertas.append({"cat": "REH", "descricao": "REH nao encontrada para esta data/subgrupo", "status": "INVESTIGAR"})

    # ----------------------------------------------------------------
    # 2. Verificar tarifas cobradas vs REH
    # ----------------------------------------------------------------
    consumo_kwh = dados.get("consumo_faturado") or 0
    itens = dados.get("itens", [])

    tusd_item = next((i for i in itens if i.get("tipo") == "consumo_tusd"), None)
    te_item = next((i for i in itens if i.get("tipo") == "consumo_te"), None)

    if tusd_item and posto_reh:
        tar_cob = tusd_item.get("preco_unit_com_trib") or 0
        cobrado["tusd_com_trib"] = tar_cob
        diff_pct = abs(tar_cob - tusd_com) / tusd_com if tusd_com else 0
        status = "OK" if diff_pct <= tol_pct else "INVESTIGAR"
        alertas.append({
            "cat": "Tarifa TUSD",
            "descricao": f"TUSD cobrada R${tar_cob:.6f}/kWh vs auditada R${tusd_com:.6f}/kWh",
            "status": status,
            "diferenca": round(tar_cob - tusd_com, 6),
        })

    if te_item and posto_reh:
        tar_cob = te_item.get("preco_unit_com_trib") or 0
        cobrado["te_com_trib"] = tar_cob
        diff_pct = abs(tar_cob - te_com) / te_com if te_com else 0
        status = "OK" if diff_pct <= tol_pct else "INVESTIGAR"
        alertas.append({
            "cat": "Tarifa TE",
            "descricao": f"TE cobrada R${tar_cob:.6f}/kWh vs auditada R${te_com:.6f}/kWh",
            "status": status,
            "diferenca": round(tar_cob - te_com, 6),
        })

    # ----------------------------------------------------------------
    # 3. Bandeira proporcional
    # ----------------------------------------------------------------
    bandeira_tab = find_bandeira(data_ref, bandeiras)
    dias_ciclo = dados.get("dias_ciclo") or 30
    dias_patamar = dados.get("dias_por_patamar") or {}

    if not dias_patamar:
        bv = dados.get("bandeira_vigente", "VERDE")
        dias_patamar = {bv: dias_ciclo}

    band_auditada = 0.0
    dias_estimados = {}
    linhas_detalhadas = []

    for pat, dias in dias_patamar.items():
        tar_band_sem = bandeira_tab.get(pat, 0)
        kwh_pat = round((dias / dias_ciclo) * consumo_kwh, 3)
        valor_sem = kwh_pat * tar_band_sem
        valor_com = grossup(valor_sem, icms_alq, pis_alq, cofins_alq) if valor_sem > 0 else 0
        valor_pat = round(valor_com, 2)
        band_auditada += valor_pat
        dias_estimados[pat] = dias
        if tar_band_sem > 0:
            linhas_detalhadas.append(
                f"{pat}: ({dias}/{dias_ciclo} dias) x {consumo_kwh:.0f} kWh "
                f"= {kwh_pat:.3f} kWh x R${tar_band_sem:.5f} (sem trib) "
                f"+ grossup = R${valor_pat:.2f}"
            )
        else:
            linhas_detalhadas.append(
                f"{pat}: ({dias}/{dias_ciclo} dias) x {consumo_kwh:.0f} kWh "
                f"= {kwh_pat:.3f} kWh x R$0,00 = R$0,00"
            )
    band_auditada = round(band_auditada, 2)
    auditado["bandeira_auditada"] = band_auditada
    auditado["dias_estimados_por_patamar"] = dias_estimados

    band_cobrada = sum(i.get("valor") or 0 for i in itens if i.get("tipo") == "bandeira")
    cobrado["bandeira"] = band_cobrada
    diff_band = abs(band_auditada - band_cobrada)

    detalhes_str = " | ".join(linhas_detalhadas) if linhas_detalhadas else "Bandeira VERDE (sem adicional)"
    alertas.append({
        "cat": "Bandeira",
        "descricao": (
            f"{detalhes_str} -> Total auditado R${band_auditada:.2f} | Cobrado R${band_cobrada:.2f}"
        ),
        "status": "OK" if diff_band <= tol_rs else "INVESTIGAR",
        "diferenca": round(band_cobrada - band_auditada, 2),
    })

    # ----------------------------------------------------------------
    # 4. Tributos - Tese do Seculo (PIS/COFINS base = Base ICMS - ICMS)
    # ----------------------------------------------------------------
    icms_val_cob = trib.get("icms", {}).get("valor") or 0
    base_icms_cob = trib.get("icms", {}).get("base") or 0
    base_pis_cob = trib.get("pis", {}).get("base") or 0
    base_pis_aud = base_icms_cob - icms_val_cob
    diff_base_pis = abs(base_pis_aud - base_pis_cob)
    alertas.append({
        "cat": "Tese do Seculo",
        "descricao": f"Base PIS/COFINS cobrada R${base_pis_cob:.2f} vs auditada (Base ICMS - ICMS) R${base_pis_aud:.2f}",
        "status": "OK" if diff_base_pis <= tol_rs else "INVESTIGAR",
        "diferenca": round(base_pis_cob - base_pis_aud, 2),
    })

    # ----------------------------------------------------------------
    # 5. Periodo de leitura (REN 1.000/2021: 15-45 dias)
    # ----------------------------------------------------------------
    dias_min = params.get("ciclo_dias_min", 15)
    dias_max = params.get("ciclo_dias_max", 45)
    if dias_ciclo:
        status_dias = "OK" if dias_min <= dias_ciclo <= dias_max else "ATENCAO"
        alertas.append({
            "cat": "Periodo Leitura",
            "descricao": f"Ciclo de {dias_ciclo} dias (REN 1.000/2021: {dias_min}-{dias_max} dias)",
            "status": status_dias,
            "valor": dias_ciclo,
        })

    # ----------------------------------------------------------------
    # 6. Fio B (Lei 14.300/2022)
    # ----------------------------------------------------------------
    data_adesao = config.get("data_adesao_mmgd", "2022-01-01")
    fator_fio_b = calcular_fio_b(data_ref, data_adesao, fio_b_data)
    auditado["fator_fio_b"] = fator_fio_b

    if fator_fio_b == 0:
        alertas.append({
            "cat": "Fio B",
            "descricao": "Sistema pre-MMGD: isento de Fio B ate 2045 (Art. 26 Lei 14.300/2022)",
            "status": "OK",
        })
    else:
        alertas.append({
            "cat": "Fio B",
            "descricao": f"Sistema pos-MMGD: fator Fio B = {fator_fio_b*100:.0f}% do componente TUSD-Fio B",
            "status": "ATENCAO",
            "valor": fator_fio_b,
        })

    # ----------------------------------------------------------------
    # 7. Cobranças retroativas
    # ----------------------------------------------------------------
    tipos_retro = ["juros_mora", "multa", "atualizacao_monetaria", "icms_cde"]
    retro_items = [i for i in itens if i.get("tipo") in tipos_retro]
    if retro_items:
        total_retro = sum(i.get("valor") or 0 for i in retro_items)
        alertas.append({
            "cat": "Retroativo",
            "descricao": f"Cobrancas retroativas detectadas: {[i['tipo'] for i in retro_items]} - Total R${total_retro:.2f}",
            "status": "INVESTIGAR",
            "valor_total": round(total_retro, 2),
        })

    # ----------------------------------------------------------------
    # 8. Total a pagar vs soma dos itens
    # ----------------------------------------------------------------
    total_fatura = dados.get("total_fatura") or dados.get("total_a_pagar") or 0
    total_a_pagar = dados.get("total_a_pagar") or 0
    diff_total = abs(total_fatura - total_a_pagar)
    if total_fatura and total_a_pagar:
        alertas.append({
            "cat": "Total a Pagar",
            "descricao": f"Total fatura R${total_fatura:.2f} vs total a pagar R${total_a_pagar:.2f}",
            "status": "OK" if diff_total <= tol_rs else "INVESTIGAR",
            "diferenca": round(total_fatura - total_a_pagar, 2),
        })

    # ----------------------------------------------------------------
    # 9. GD: compensacao > injetada
    # ----------------------------------------------------------------
    if config.get("tem_gd"):
        energia_comp = config.get("energia_compensada_kwh") or 0
        injetada = dados.get("gd_injetada_mes") or 0
        if energia_comp and injetada and energia_comp > injetada * 1.01:
            alertas.append({
                "cat": "GD Compensacao",
                "descricao": f"Compensacao ({energia_comp} kWh) maior que injetada no mes ({injetada} kWh) - uso de saldo ou geracao compartilhada",
                "status": "INVESTIGAR",
                "compensada": energia_comp,
                "injetada": injetada,
            })
        else:
            alertas.append({
                "cat": "GD Compensacao",
                "descricao": "Compensacao dentro do esperado",
                "status": "OK",
            })

    # ----------------------------------------------------------------
    # 10. Consumo medido
    # ----------------------------------------------------------------
    medidores = dados.get("medidores") or []
    med = medidores[0] if medidores else {}
    leit_atual_num = med.get("leitura_atual")
    leit_ant_num   = med.get("leitura_anterior")
    constante      = med.get("constante") or 1.0
    consumo_fat    = dados.get("consumo_faturado") or 0

    _tipo_forn = (dados.get("tipo_fornecimento") or "").lower()
    if "trifasico" in _tipo_forn or "trifasico" in _tipo_forn:
        _minimo_cd = 100
    elif "bifasico" in _tipo_forn or "bifasico" in _tipo_forn:
        _minimo_cd = 30 if ("dois" in _tipo_forn or "2 " in _tipo_forn) else 50
    elif "monofasico" in _tipo_forn or "monofasico" in _tipo_forn:
        _minimo_cd = 30
    else:
        _minimo_cd = None

    if leit_atual_num is not None and leit_ant_num is not None and consumo_fat:
        consumo_calc      = round((leit_atual_num - leit_ant_num) * constante, 1)
        consumo_fat_arred = round(consumo_fat, 1)
        diff_kwh          = round(consumo_calc - consumo_fat_arred, 1)
        auditado["consumo_calculado"] = consumo_calc

        if _minimo_cd is not None and consumo_calc < _minimo_cd:
            if consumo_fat_arred == _minimo_cd:
                alertas.append({
                    "cat": "Consumo Medidor",
                    "descricao": (
                        f"({leit_atual_num:.0f} - {leit_ant_num:.0f}) x {constante:.2f} "
                        f"= {consumo_calc:.0f} kWh medido < minimo {_minimo_cd} kWh "
                        f"({dados.get('tipo_fornecimento', '')}) - "
                        f"Custo de Disponibilidade aplicado corretamente (Art. 291 REN 1.000/2021)"
                    ),
                    "status": "OK",
                    "diferenca": 0,
                })
            else:
                alertas.append({
                    "cat": "Consumo Medidor",
                    "descricao": (
                        f"({leit_atual_num:.0f} - {leit_ant_num:.0f}) x {constante:.2f} "
                        f"= {consumo_calc:.0f} kWh medido < minimo {_minimo_cd} kWh "
                        f"({dados.get('tipo_fornecimento', '')}) - "
                        f"faturado {consumo_fat_arred:.0f} kWh, esperado {_minimo_cd} kWh "
                        f"(Art. 291 REN 1.000/2021)"
                    ),
                    "status": "INVESTIGAR",
                    "diferenca": round(consumo_fat_arred - _minimo_cd, 1),
                })
        else:
            alertas.append({
                "cat": "Consumo Medidor",
                "descricao": (
                    f"({leit_atual_num:.0f} - {leit_ant_num:.0f}) x constante {constante:.2f} "
                    f"= {consumo_calc:.0f} kWh auditado vs {consumo_fat_arred:.0f} kWh faturado (quantidade)"
                ),
                "status": "OK" if diff_kwh == 0 else "INVESTIGAR",
                "diferenca": diff_kwh,
            })

    # ----------------------------------------------------------------
    # 11. Impedimento / estimativa de leitura
    # ----------------------------------------------------------------
    if dados.get("leitura_estimada"):
        alertas.append({
            "cat": "Tipo de Leitura",
            "descricao": (
                f"Aviso nas observacoes da fatura: '{dados.get('leitura_aviso')}' - "
                "o consumo pode ser estimado e nao medido; verifique as leituras reais"
            ),
            "status": "INVESTIGAR",
        })

    # ----------------------------------------------------------------
    # 12. Variacao de consumo vs historico
    # ----------------------------------------------------------------
    historico = dados.get("historico_consumo") or []
    if historico and consumo_fat:
        mes_atual_abrev = (dados.get("mes_ref") or "")[:3].upper()
        hist_prev = [
            h for h in historico
            if h.get("consumo_kwh") and h.get("mes", "").upper() != mes_atual_abrev
        ]

        if hist_prev:
            h_ant = hist_prev[-1]
            cons_ant = h_ant.get("consumo_kwh", 0)
            mes_ant_nome = h_ant.get("mes", "mes anterior")
            if cons_ant and cons_ant > 0:
                var_pct = (consumo_fat - cons_ant) / cons_ant * 100
                if var_pct > 30:
                    st_var = "INVESTIGAR"
                elif var_pct > 15:
                    st_var = "ATENCAO"
                else:
                    st_var = "OK"
                alertas.append({
                    "cat": "Variacao Mensal",
                    "descricao": (
                        f"Consumo atual {consumo_fat:.0f} kWh vs {mes_ant_nome} "
                        f"{cons_ant} kWh ({var_pct:+.1f}%)"
                    ),
                    "status": st_var,
                    "diferenca": round(var_pct, 1),
                })

            if len(hist_prev) >= 2:
                consumos_hist = [h["consumo_kwh"] for h in hist_prev]
                media_hist = sum(consumos_hist) / len(consumos_hist)
                var_media_pct = (consumo_fat - media_hist) / media_hist * 100
                if var_media_pct > 40:
                    st_media = "INVESTIGAR"
                elif var_media_pct > 20:
                    st_media = "ATENCAO"
                else:
                    st_media = "OK"
                alertas.append({
                    "cat": "Variacao Historica",
                    "descricao": (
                        f"Consumo atual {consumo_fat:.0f} kWh vs media historica "
                        f"{media_hist:.0f} kWh ({var_media_pct:+.1f}%) "
                        f"- base: {len(consumos_hist)} meses"
                    ),
                    "status": st_media,
                    "diferenca": round(var_media_pct, 1),
                })

    return {
        "auditado": auditado,
        "cobrado": cobrado,
        "alertas": alertas,
    }


if __name__ == "__main__":
    import sys
    sys.path.insert(0, str(Path(__file__).parent))
    from extractor import extract_fatura
    if len(sys.argv) < 2:
        print("Uso: python audit.py <fatura.pdf>")
        sys.exit(1)
    dados = extract_fatura(sys.argv[1])
    config = {
        "data_adesao_mmgd": "2022-01-01",
        "tem_gd": dados.get("tem_gd", False),
        "energia_compensada_kwh": dados.get("gd_ajuste_cat") or dados.get("gd_injetada_mes"),
    }
    result = auditar_fatura(dados, config)
    for a in result["alertas"]:
        icon = {"OK": "OK", "ATENCAO": "(!)", "INVESTIGAR": "(?)"}.get(a["status"], "?")
        print(f"{icon} [{a['cat']}] {a['descricao']}")

# -*- coding: utf-8 -*-
"""
audit.py - Logica de triagem para faturas Neoenergia PE
Faixas: OK / INVESTIGAR / DIVERGENCIA
"""
import unicodedata


def _norm(s):
    """Remove acentos e coloca em lower para comparacoes de tipo_fornecimento."""
    return unicodedata.normalize("NFD", s.lower()).encode("ascii", "ignore").decode()


TOL_ITEM   = 0.10
TOL_ICMS   = 0.10
TOL_LEIT   = 2
TOL_SOMA_OK   = 0.10
TOL_SOMA_INV  = 0.30

CAMPOS_CRIT = [
    "ref_mes_ano", "vencimento", "conta_contrato",
    "consumo_kwh_tusd_qtd", "valor_tusd", "valor_te",
    "preco_tusd", "preco_te", "total_fatura",
]
CAMPOS_SEC = [
    "cod_instalacao", "data_emissao", "nr_medidor",
    "leitura_anterior", "leitura_atual", "icms_aliq",
]


def _chk(val, tol):
    if val is None:
        return None
    return abs(val) <= tol


def _audit_leitura(lant, latu, cte, qtd_total, label, flags_inv, metricas, tol=TOL_LEIT):
    """Audita (latu - lant) * cte para um medidor com leitura_anterior conhecida (> 0)."""
    calc = round((latu - lant) * cte, 1)
    metricas[f"calc_leit_kWh_{label}"] = calc
    if calc < 0:
        flags_inv.append(f"leitura {label}: leitura_atual {latu} < leitura_anterior {lant}")
    elif calc > qtd_total + tol:
        flags_inv.append(
            f"leitura {label}: ({latu}-{lant})x{cte}={calc:.1f}kWh "
            f"> consumo_total {qtd_total:.2f}kWh"
        )
    return calc


def auditar(r):
    motivos  = []
    metricas = {}
    flags_div  = []
    flags_inv  = []

    # 1. Campos criticos ausentes
    faltam_crit = [c for c in CAMPOS_CRIT if r.get(c) is None]
    faltam_sec  = [c for c in CAMPOS_SEC  if r.get(c) is None]
    if faltam_crit:
        flags_div.append(f"campos criticos ausentes: {', '.join(faltam_crit)}")
    elif faltam_sec:
        flags_inv.append(f"campos secundarios ausentes: {', '.join(faltam_sec)}")

    # 2. Math TUSD
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

    # 3. Math TE
    pte  = r.get("preco_te")
    vte  = r.get("valor_te")
    qtde = r.get("consumo_kwh_te_qtd") or qtd
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

    # 4. Leitura do medidor
    cte = r.get("constante_medidor") or 1.0
    if r.get("nr_medidores", 1) > 1:
        # Troca de medidor no ciclo: auditar cada medidor com leitura_anterior > 0
        metricas["troca_medidor"] = True
        soma_auditada = 0.0
        n_auditados   = 0
        for sufixo, label in [("", "med1"), ("_2", "med2")]:
            lant = r.get(f"leitura_anterior{sufixo}")
            latu = r.get(f"leitura_atual{sufixo}")
            if lant is not None and latu is not None and lant > 0 and qtd is not None:
                calc = _audit_leitura(lant, latu, cte, qtd, label, flags_inv, metricas)
                soma_auditada += calc
                n_auditados   += 1
            else:
                metricas[f"calc_leit_kWh_{label}"] = None  # lant=0 ou ausente, nao auditavel
        metricas["soma_leit_auditada_kWh"] = round(soma_auditada, 1) if n_auditados else None
        metricas["dif_leit_kWh"] = None  # sem auditoria direta de diferenca total
    else:
        lant = r.get("leitura_anterior")
        latu = r.get("leitura_atual")
        if lant is not None and latu is not None and qtd is not None:
            calc_leit = round((latu - lant) * cte, 1)
            dif_leit  = abs(calc_leit - qtd)
            metricas["calc_leit_kWh"] = calc_leit
            metricas["dif_leit_kWh"]  = round(dif_leit, 1)
            if dif_leit > TOL_LEIT:
                # Verificar se a diferenca e a compensacao GDI com deducao direta
                # (REN 1000/2021 / Lei 14.300/2022): distribuidora subtrai kWh compensados
                # do consumo medido antes de faturar, sem linha de credito separada.
                scee_kwh_gdi = (r.get("scee_kwh_compensados") or 0) if r.get("is_scee") else 0
                comp_c_gdi   = r.get("valor_imp_som_dim_c") or 0
                if (scee_kwh_gdi > 0
                        and abs(comp_c_gdi) < TOL_ITEM
                        and abs(dif_leit - scee_kwh_gdi) <= TOL_LEIT):
                    metricas["gdi_deducao_direta"]      = True
                    metricas["gdi_scee_kwh_verificado"] = round(dif_leit, 2)
                else:
                    _tipo = _norm(r.get("tipo_fornecimento") or "")
                    if "trifasico" in _tipo:
                        _minimo = 100
                    elif "bifasico" in _tipo:
                        _minimo = 30 if ("dois" in _tipo or "2 " in _tipo) else 50
                    elif "monofasico" in _tipo:
                        _minimo = 30
                    else:
                        _minimo = None

                    if _minimo is not None and calc_leit < _minimo and round(qtd, 1) == _minimo:
                        metricas["custo_disponibilidade"] = _minimo
                    else:
                        flags_inv.append(
                            f"leitura: ({latu}-{lant})x{cte}={calc_leit} "
                            f"!= consumo={qtd} (dif={dif_leit:.1f}kWh)"
                        )
        else:
            metricas["dif_leit_kWh"] = None

    # 5. Math ICMS
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

    # 5b. SCEE -- auditoria da compensacao
    if r.get("is_scee"):
        scee_kwh  = r.get("scee_kwh_compensados") or 0
        comp_c    = r.get("valor_imp_som_dim_c") or 0
        preco_tot = (r.get("preco_tusd") or 0) + (r.get("preco_te") or 0)
        metricas["scee_kwh_compensados"] = scee_kwh
        metricas["is_scee"] = True
        if metricas.get("gdi_deducao_direta"):
            # GDI com deducao direta (REN 1000/2021 / Lei 14.300/2022):
            # Math verificado: consumo_faturado = consumo_medido - scee_kwh_compensados (sec.4)
            # TUSD e TE auditados sobre consumo_faturado (secs. 2 e 3).
            # Valor implicito da compensacao registrado apenas como metrica.
            metricas["gdi_comp_monetario_equiv_R$"] = round(scee_kwh * preco_tot, 2)
        elif scee_kwh == 0:
            pass
        elif scee_kwh and preco_tot:
            comp_aud = round(scee_kwh * preco_tot, 2)
            comp_cob = round(abs(comp_c), 2)
            dif_scee = abs(comp_aud - comp_cob)
            metricas["comp_scee_auditado_R$"] = comp_aud
            metricas["comp_scee_cobrado_R$"]  = comp_cob
            metricas["dif_scee_R$"]           = round(dif_scee, 2)
            if dif_scee > TOL_ITEM:
                flags_inv.append(
                    f"SCEE: {scee_kwh}kWh x R${preco_tot:.6f}/kWh "
                    f"= R${comp_aud:.2f} auditado vs R${comp_cob:.2f} cobrado "
                    f"(dif=R${dif_scee:.2f})"
                )
        else:
            flags_inv.append("SCEE: kWh compensados presentes mas preco_tusd/preco_te ausentes")

    # 6. Soma dos itens vs total da fatura
    total = r.get("total_fatura")
    itens = [
        r.get("valor_tusd"),
        r.get("valor_te"),
        r.get("valor_bandeira") or 0,
        r.get("cosip") or 0,
        r.get("icms_cde") or 0,
        r.get("valor_parcelamento") or 0,
        r.get("valor_ipca") or 0,
        r.get("valor_imp_som_dim_c") or 0,
        r.get("valor_imp_som_dim_s") or 0,
        r.get("valor_religacao") or 0,
        r.get("valor_multas_nf") or 0,
        r.get("valor_juros_nf") or 0,
        r.get("valor_encargos_cosip") or 0,
    ]
    soma = sum(v for v in itens if v is not None)
    metricas["soma_itens_R$"] = round(soma, 2)
    if total and soma:
        dif_total = soma - total
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

    # 7. Erros de extracao
    if r.get("erro"):
        flags_div.append(f"erro de extracao: {r['erro']}")

    # Triagem final
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


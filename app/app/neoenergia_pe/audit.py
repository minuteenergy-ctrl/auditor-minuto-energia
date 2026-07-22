# -*- coding: utf-8 -*-
"""
audit.py - Logica de triagem para faturas Neoenergia PE
Faixas: OK / INVESTIGAR / DIVERGENCIA
Suporta BT (Grupo B) e MT (Grupo A, Tarifa A4 Azul).
"""
import json
import unicodedata
from datetime import datetime, timedelta
from pathlib import Path

_DATA_DIR = Path(__file__).parent / "data"
_TARIFAS = None


def _carregar_tarifas():
    global _TARIFAS
    if _TARIFAS is None:
        with open(_DATA_DIR / "tarifas_neo_pe.json", encoding="utf-8") as f:
            _TARIFAS = json.load(f)
    return _TARIFAS


def _parse_ref(ref_mes_ano):
    """
    Converte 'MM/AAAA' em datetime(ano, mes, 1). Retorna None se invalido.
    As datas de vigencia no JSON sao billing-cycle-aligned (1o dia do mes),
    portanto a comparacao pelo 1o dia do mes e suficiente e correta.
    """
    if not ref_mes_ano:
        return None
    partes = str(ref_mes_ano).split("/")
    if len(partes) != 2:
        return None
    try:
        return datetime(int(partes[1]), int(partes[0]), 1)
    except (ValueError, TypeError):
        return None


def _reh_para_lista(ref_mes_ano, lista_rehs):
    """Retorna a REH vigente da lista para ref_mes_ano ('MM/AAAA'), ou None."""
    ref_dt = _parse_ref(ref_mes_ano)
    if ref_dt is None:
        return None
    for t in lista_rehs:
        ini = datetime.strptime(t["vigencia_inicio"], "%Y-%m-%d")
        fim_str = t.get("vigencia_fim")
        fim = datetime.strptime(fim_str, "%Y-%m-%d") if fim_str else datetime(9999, 12, 31)
        if ini <= ref_dt <= fim:
            return t
    return None


def _reh_bt_para_periodo(ref_mes_ano):
    """Retorna dict REH BT vigente para ref_mes_ano ('MM/AAAA'), ou None se nao cadastrado."""
    tarifas = _carregar_tarifas()
    return _reh_para_lista(ref_mes_ano, tarifas.get("tarifas_bt", []))


def _reh_bt_ponderado(data_lant_str, data_latu_str, lista_rehs):
    """
    Calcula tarifas BT ponderadas pelos dias sob cada REH no periodo de leitura.
    Usa data_oficial (se disponivel no JSON) ou vigencia_inicio como inicio efetivo.
    Periodo de consumo: dia seguinte a leitura_anterior ate leitura_atual (inclusive).
    Retorna dict {TUSD_kwh, TE_kwh, reh, ponderado} ou None se datas invalidas.
    """
    try:
        dt_ant = datetime.strptime(data_lant_str, "%d/%m/%Y")
        dt_atu = datetime.strptime(data_latu_str, "%d/%m/%Y")
    except (ValueError, TypeError):
        return None

    total_dias = (dt_atu - dt_ant).days   # = "N DE DIAS" da fatura
    if total_dias <= 0:
        return None

    dt_ini = dt_ant + timedelta(days=1)   # primeiro dia de consumo

    # Ordena REHs cronologicamente para calcular quando cada uma termina
    rehs_ord = sorted(lista_rehs, key=lambda t: t["vigencia_inicio"])

    tusd_pond = te_pond = 0.0
    reh_labels = []

    for i, t in enumerate(rehs_ord):
        # Inicio efetivo: data_oficial (data ANEEL real) se disponivel, senao vigencia_inicio
        ini_str = t.get("data_oficial") or t["vigencia_inicio"]
        ini = datetime.strptime(ini_str, "%Y-%m-%d")

        # Fim exclusivo: inicio efetivo da proxima REH
        if i + 1 < len(rehs_ord):
            prox_str = rehs_ord[i + 1].get("data_oficial") or rehs_ord[i + 1]["vigencia_inicio"]
            fim_exc = datetime.strptime(prox_str, "%Y-%m-%d")
        else:
            fim_exc = datetime(9999, 12, 31)

        # Overlap entre periodo de consumo [dt_ini, dt_atu+1) e [ini, fim_exc)
        ov_ini = max(dt_ini, ini)
        ov_fim = min(dt_atu + timedelta(days=1), fim_exc)

        if ov_fim > ov_ini:
            dias = (ov_fim - ov_ini).days
            tusd_pond += dias * t["TUSD_kwh"]
            te_pond   += dias * t["TE_kwh"]
            reh_labels.append(f"{t['reh']} ({dias}d/{total_dias}d)")

    if not reh_labels:
        return None

    return {
        "TUSD_kwh":  round(tusd_pond / total_dias, 8),
        "TE_kwh":    round(te_pond   / total_dias, 8),
        "reh":       " + ".join(reh_labels),
        "ponderado": len(reh_labels) > 1,
    }


def _reh_mt_para_periodo(ref_mes_ano, subgrupo="A4", modalidade="Azul"):
    """Retorna dict REH MT vigente para ref_mes_ano, ou None se nao cadastrado."""
    chave = f"{subgrupo}_{modalidade}"
    tarifas = _carregar_tarifas()
    return _reh_para_lista(ref_mes_ano, tarifas.get("tarifas_mt", {}).get(chave, []))


def _norm(s):
    """Remove acentos e coloca em lower para comparacoes de tipo_fornecimento."""
    return unicodedata.normalize("NFD", s.lower()).encode("ascii", "ignore").decode()


def _tol_reh(cobrado, esperado, tol=0.0005):
    """True se diferenca relativa > tol (default 0.05%)."""
    if esperado is None or esperado == 0:
        return False
    return abs(cobrado - esperado) / abs(esperado) > tol


TOL_ITEM      = 0.10
TOL_ICMS      = 0.10
TOL_LEIT      = 2
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

CAMPOS_CRIT_MT = [
    "ref_mes_ano", "vencimento", "conta_contrato", "total_fatura",
    "consumo_ponta_kwh", "consumo_fp_kwh",
    "valor_fio_np", "valor_fio_fp",
    "valor_encar_np", "valor_encar_fp",
]
CAMPOS_SEC_MT = [
    "cod_instalacao", "data_emissao",
    "dem_contratada_fp_kw", "dem_contratada_np_kw",
    "tarifa_fio_np_sem", "tarifa_fio_fp_sem",
    "tarifa_encar_np_sem",
    "icms_aliq",
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


# ─────────────────────────── AUDITORIA MT ───────────────────────────────────

def _auditar_mt(r):
    """Auditoria completa para faturas MT (Grupo A) Neoenergia PE."""
    flags_div = []
    flags_inv = []
    metricas  = {}

    # 1. Campos criticos ausentes
    faltam_crit = [c for c in CAMPOS_CRIT_MT if r.get(c) is None]
    faltam_sec  = [c for c in CAMPOS_SEC_MT  if r.get(c) is None]
    if faltam_crit:
        flags_div.append(f"campos criticos ausentes: {', '.join(faltam_crit)}")
    elif faltam_sec:
        flags_inv.append(f"campos secundarios ausentes: {', '.join(faltam_sec)}")

    # 2. REH MT — validacao de tarifas sem tributos
    subgrupo   = r.get("subgrupo",   "A4")
    modalidade = r.get("modalidade", "Azul")
    reh_mt = _reh_mt_para_periodo(r.get("ref_mes_ano"), subgrupo, modalidade)

    if reh_mt is None:
        flags_inv.append(
            f"REH MT {subgrupo} {modalidade}: nenhuma REH cadastrada para "
            f"periodo {r.get('ref_mes_ano')} -- verificar manualmente"
        )
    else:
        metricas["reh_mt"] = reh_mt.get("reh")
        checks_mt = [
            ("tarifa Fio NP",  r.get("tarifa_fio_np_sem"),  reh_mt.get("Fio_NP_kw"),         "R$/kW"),
            ("tarifa Fio FP",  r.get("tarifa_fio_fp_sem"),  reh_mt.get("Fio_FP_kw"),         "R$/kW"),
            ("tarifa Encar NP",r.get("tarifa_encar_np_sem"),reh_mt.get("Encar_NP_kwh"),      "R$/kWh"),
            ("tarifa Encar FP",r.get("tarifa_encar_fp_sem"),reh_mt.get("Encar_FP_kwh"),      "R$/kWh"),
        ]
        for nome, cobrado, esperado, unidade in checks_mt:
            if cobrado is None:
                flags_inv.append(f"REH MT {nome}: nao extraida da fatura -- verificar manualmente")
                continue
            if esperado is None:
                continue
            if _tol_reh(cobrado, esperado):
                flags_div.append(
                    f"REH MT {nome}: fatura={cobrado:.8f} vs "
                    f"{reh_mt['reh']}={esperado:.8f} {unidade} "
                    f"(dif={cobrado - esperado:+.8f})"
                )

    # 3. Math: Fio NP (dem × preco_com = valor)
    for tag, key_qtd, key_preco, key_valor in [
        ("Fio NP",    "dem_fio_np_kw",       "preco_fio_np_com",  "valor_fio_np"),
        ("Fio FP",    "dem_fio_fp_kw",       "preco_fio_fp_com",  "valor_fio_fp"),
        ("Encar NP",  "consumo_encar_np_kwh","preco_encar_np_com","valor_encar_np"),
        ("Encar FP",  "consumo_encar_fp_kwh","preco_encar_fp_com","valor_encar_fp"),
    ]:
        qtd   = r.get(key_qtd)
        preco = r.get(key_preco)
        valor = r.get(key_valor)
        if qtd is not None and preco is not None and valor is not None:
            calc = round(qtd * preco, 2)
            dif  = abs(calc - valor)
            metricas[f"calc_{tag.replace(' ','_')}_R$"] = calc
            if dif > TOL_ITEM:
                flags_div.append(
                    f"{tag} math: {qtd} x {preco} = {calc:.2f} != {valor:.2f} "
                    f"(dif={dif:.2f})"
                )
        elif valor is not None:
            flags_inv.append(f"{tag}: qtd/preco nao extraidos -- math nao auditavel")

    # 4. Consistencia: qtd Encar == consumo do Demonstrativo
    for key_encar, key_dem, label in [
        ("consumo_encar_np_kwh", "consumo_ponta_kwh", "kWh NP (encar vs demonstrativo)"),
        ("consumo_encar_fp_kwh", "consumo_fp_kwh",    "kWh FP (encar vs demonstrativo)"),
    ]:
        e = r.get(key_encar)
        d = r.get(key_dem)
        if e is not None and d is not None:
            dif = abs(e - d)
            if dif > TOL_LEIT:
                flags_inv.append(
                    f"consumo {label}: item={e:.2f} vs demonstrativo={d:.2f} "
                    f"(dif={dif:.2f} kWh)"
                )

    # 5. Demanda: medida vs contratada (ultrapassagem)
    for posto, key_med, key_cont in [
        ("NP", "demanda_medida_np_kw", "dem_contratada_np_kw"),
        ("FP", "demanda_medida_fp_kw", "dem_contratada_fp_kw"),
    ]:
        med  = r.get(key_med)
        cont = r.get(key_cont)
        if med is not None and cont is not None and cont > 0:
            if med > cont * 1.05:   # ultrapassagem > 5%
                pct = (med - cont) / cont * 100
                metricas[f"ultrap_{posto}_pct"] = round(pct, 1)
                flags_inv.append(
                    f"demanda {posto}: medida={med:.2f} kW > contratada={cont:.0f} kW "
                    f"({pct:.1f}% acima) -- verificar ultrapassagem"
                )

    # 6. ICMS math
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
                f"ICMS math: {icms_b} x {icms_a}% = {calc_icms:.2f} != {icms_v:.2f} "
                f"(dif={dif_icms:.2f})"
            )
    else:
        metricas["dif_ICMS_R$"] = None

    # 7. Soma dos itens vs total da fatura
    total = r.get("total_fatura")
    itens_mt = [
        r.get("valor_fio_np")       or 0,
        r.get("valor_fio_fp")       or 0,
        r.get("valor_encar_np")     or 0,
        r.get("valor_encar_fp")     or 0,
        r.get("valor_dem_reat_np")  or 0,
        r.get("valor_dem_reat_fp")  or 0,
        r.get("valor_cons_reat_np") or 0,
        r.get("valor_cons_reat_fp") or 0,
        r.get("cosip")              or 0,
        r.get("icms_cde")           or 0,
        r.get("valor_encargos_cosip") or 0,
        r.get("valor_multas_nf")    or 0,
        r.get("valor_juros_nf")     or 0,
        r.get("valor_ipca")         or 0,
        r.get("valor_dif_desc_np")  or 0,   # negativo
        r.get("valor_dif_desc_fp")  or 0,   # negativo
        r.get("valor_imp_som_dim_mt") or 0, # positivo em MT
        r.get("valor_parcelamento") or 0,
        r.get("valor_religacao")    or 0,
    ]
    soma = round(sum(v for v in itens_mt if v is not None), 2)
    metricas["soma_itens_MT_R$"] = soma
    if total and soma:
        dif_total = soma - total
        dif_pct   = abs(dif_total) / total if total else 0
        metricas["dif_total_MT_R$"] = round(dif_total, 2)
        metricas["dif_total_MT_%"]  = round(dif_pct * 100, 1)
        if dif_pct > TOL_SOMA_INV:
            flags_div.append(
                f"soma itens R${soma:.2f} != total R${total:.2f} "
                f"(dif={dif_total:+.2f}, {dif_pct*100:.0f}%) "
                f"-- item MT nao extraido"
            )
        elif dif_pct > TOL_SOMA_OK:
            flags_inv.append(
                f"soma itens R${soma:.2f} != total R${total:.2f} "
                f"(dif={dif_total:+.2f}, {dif_pct*100:.0f}%) "
                f"-- possivel item nao extraido"
            )
    else:
        metricas["dif_total_MT_R$"] = None
        metricas["dif_total_MT_%"]  = None

    # 8. Erros de extracao
    if r.get("erro"):
        flags_div.append(f"erro de extracao: {r['erro']}")

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


# ─────────────────────────── AUDITORIA BT ───────────────────────────────────

def auditar(r):
    # Faturas MT (Grupo A): rota separada
    if r.get("is_mt"):
        return _auditar_mt(r)

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
                metricas[f"calc_leit_kWh_{label}"] = None
        metricas["soma_leit_auditada_kWh"] = round(soma_auditada, 1) if n_auditados else None
        metricas["dif_leit_kWh"] = None
    else:
        lant = r.get("leitura_anterior")
        latu = r.get("leitura_atual")
        if lant is not None and latu is not None and qtd is not None:
            if latu < lant:
                _capacidade = 10 ** len(str(int(lant)))
                if lant >= _capacidade * 0.8:
                    calc_leit = round((_capacidade - lant + latu) * cte, 1)
                    metricas["virada_medidor"]     = True
                    metricas["capacidade_medidor"] = int(_capacidade)
                else:
                    metricas["virada_medidor"]     = True
                    metricas["capacidade_medidor"] = None
                    flags_inv.append(
                        f"leitura: virada de medidor (latu={latu} < lant={lant}) "
                        f"-- capacidade nao determinavel, verificar manualmente"
                    )
                    calc_leit = None
            else:
                calc_leit = round((latu - lant) * cte, 1)

            if calc_leit is not None:
                dif_leit  = abs(calc_leit - qtd)
                metricas["calc_leit_kWh"] = calc_leit
                metricas["dif_leit_kWh"]  = round(dif_leit, 1)
                if dif_leit > TOL_LEIT:
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
                metricas["calc_leit_kWh"] = None
                metricas["dif_leit_kWh"]  = None
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

    # 7. REH BT — validacao de tarifas sem tributos
    #    Usa ponderacao proporcional pelos dias sob cada REH no periodo de leitura
    #    quando data_leitura_anterior e data_leitura_atual estiverem disponiveis.
    lista_bt = _carregar_tarifas().get("tarifas_bt", [])
    reh_bt = _reh_bt_ponderado(
        r.get("data_leitura_anterior"), r.get("data_leitura_atual"), lista_bt
    )
    if reh_bt is None:
        # Fallback: lookup simples por ref_mes_ano (sem datas de leitura)
        reh_bt = _reh_bt_para_periodo(r.get("ref_mes_ano"))

    if reh_bt is None:
        flags_inv.append(
            f"REH BT: nenhuma REH cadastrada para periodo {r.get('ref_mes_ano')} "
            f"-- verificar manualmente"
        )
    else:
        metricas["reh_bt"] = reh_bt.get("reh")
        if reh_bt.get("ponderado"):
            metricas["reh_bt_ponderado"] = True
        tusd_sem = r.get("tarifa_tusd_sem_trib")
        te_sem   = r.get("tarifa_te_sem_trib")
        if tusd_sem is not None:
            if _tol_reh(tusd_sem, reh_bt["TUSD_kwh"]):
                flags_div.append(
                    f"TUSD tarifa REH: fatura={tusd_sem:.8f} vs "
                    f"{reh_bt['reh']}={reh_bt['TUSD_kwh']:.8f} R$/kWh "
                    f"(dif={tusd_sem - reh_bt['TUSD_kwh']:+.8f})"
                )
        else:
            flags_inv.append(
                "TUSD: tarifa sem tributos nao extraida -- verificar REH manualmente"
            )
        if te_sem is not None:
            if _tol_reh(te_sem, reh_bt["TE_kwh"]):
                flags_div.append(
                    f"TE tarifa REH: fatura={te_sem:.8f} vs "
                    f"{reh_bt['reh']}={reh_bt['TE_kwh']:.8f} R$/kWh "
                    f"(dif={te_sem - reh_bt['TE_kwh']:+.8f})"
                )
        else:
            flags_inv.append(
                "TE: tarifa sem tributos nao extraida -- verificar REH manualmente"
            )

    # 8. Erros de extracao
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

"""
Auditor de Faturas — Minuto Energia
Interface Streamlit com identidade visual Minuto Energia
"""
import streamlit as st
import pandas as pd
import datetime
import zipfile
import io
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from extractor import extract_fatura
from audit import auditar_fatura, summary_alertas
from excel_filler import preencher_template
from excel_mestre_core import gerar_excel_mestre

# Relatório PDF
try:
    from relatorio_pdf import gerar_relatorio_pdf
    RELATORIO_PDF_DISPONIVEL = True
except ImportError:
    RELATORIO_PDF_DISPONIVEL = False

# Neoenergia PE — importação condicional
try:
    from neoenergia_pe.extractor import parse_fatura as neo_parse
    from neoenergia_pe.audit import auditar as neo_auditar
    NEO_DISPONIVEL = True
except ImportError:
    NEO_DISPONIVEL = False

# CPFL Paulista — importação condicional
try:
    from cpfl_paulista.extractor import parse_fatura as paulista_parse
    from cpfl_paulista.audit import auditar as paulista_auditar
    PAULISTA_DISPONIVEL = True
except ImportError:
    PAULISTA_DISPONIVEL = False

# ── Normalização para Excel-mestre unificado ─────────────────────────────────

def _fmt_date(v):
    if v is None:
        return None
    if hasattr(v, "strftime"):
        return v.strftime("%d/%m/%Y")
    return str(v)


def normalizar_neoenergia_pe(rec, triagem, motivos, metricas):
    return {
        "arquivo":        rec.get("arquivo"),
        "distribuidora":  "Neoenergia PE",
        "layout":         rec.get("layout"),
        "ref_mes_ano":    rec.get("ref_mes_ano"),
        "vencimento":     rec.get("vencimento"),
        "conta_uc":       rec.get("conta_contrato"),
        "cliente_nome":   None,
        "subgrupo":       "B3",
        "data_emissao":   rec.get("data_emissao"),
        "nr_nota_fiscal": rec.get("nr_nota_fiscal"),
        "nr_medidor":     rec.get("nr_medidor"),
        "consumo_kwh":    rec.get("consumo_kwh_tusd_qtd"),
        "nr_dias":        rec.get("nr_dias"),
        "leitura_anterior": rec.get("leitura_anterior"),
        "leitura_atual":    rec.get("leitura_atual"),
        "preco_tusd":     rec.get("preco_tusd"),
        "valor_tusd":     rec.get("valor_tusd"),
        "preco_te":       rec.get("preco_te"),
        "valor_te":       rec.get("valor_te"),
        "tarifa_tusd_sem": rec.get("tarifa_tusd_sem_trib"),
        "tarifa_te_sem":   rec.get("tarifa_te_sem_trib"),
        "bandeira":       rec.get("bandeira_cor"),
        "valor_bandeira": rec.get("valor_bandeira"),
        "cosip":          rec.get("cosip"),
        "total_fatura":   rec.get("total_fatura"),
        "icms_base":      rec.get("icms_base"),
        "icms_aliq":      rec.get("icms_aliq"),
        "icms_valor":     rec.get("icms_valor"),
        "pis_aliq":       rec.get("pis_aliq"),
        "pis_valor":      rec.get("pis_valor"),
        "cofins_aliq":    rec.get("cofins_aliq"),
        "cofins_valor":   rec.get("cofins_valor"),
        "__triagem__":       triagem,
        "__motivos__":       " | ".join(motivos) if motivos else "",
        "__dif_tusd__":      metricas.get("dif_TUSD_R$"),
        "__dif_te__":        metricas.get("dif_TE_R$"),
        "__dif_band__":      None,
        "__dif_leit__":      metricas.get("dif_leit_kWh"),
        "__dif_icms__":      metricas.get("dif_ICMS_R$"),
        "__dif_total__":     metricas.get("dif_total_R$"),
        "__dif_total_pct__": metricas.get("dif_total_%"),
    }


def normalizar_cpfl_paulista(rec, triagem, motivos, metricas):
    return {
        "arquivo":        rec.get("arquivo"),
        "distribuidora":  "CPFL Paulista",
        "layout":         "Verde-A4",
        "ref_mes_ano":    rec.get("ref_mes_ano"),
        "vencimento":     rec.get("vencimento"),
        "conta_uc":       rec.get("conta_contrato"),
        "cliente_nome":   rec.get("cliente_nome"),
        "subgrupo":       rec.get("subgrupo", "A4"),
        "data_emissao":   rec.get("data_emissao"),
        "nr_nota_fiscal": rec.get("nr_nota_fiscal"),
        "nr_medidor":     rec.get("nr_medidor"),
        "consumo_kwh":    (rec.get("consumo_ponta_kwh") or 0) + (rec.get("consumo_fp_kwh") or 0),
        "nr_dias":        rec.get("nr_dias"),
        "leitura_anterior": None,
        "leitura_atual":    None,
        "preco_tusd":     rec.get("tusd_fp_sem"),
        "valor_tusd":     (rec.get("valor_tusd_ponta") or 0) + (rec.get("valor_tusd_fp") or 0),
        "preco_te":       rec.get("te_fp_sem"),
        "valor_te":       (rec.get("valor_te_ponta") or 0) + (rec.get("valor_te_fp") or 0),
        "tarifa_tusd_sem": rec.get("tusd_fp_sem"),
        "tarifa_te_sem":   rec.get("te_fp_sem"),
        "bandeira":       rec.get("bandeira_cor"),
        "valor_bandeira": rec.get("valor_bandeira"),
        "cosip":          rec.get("cosip"),
        "total_fatura":   rec.get("total_fatura"),
        "icms_base":      rec.get("icms_base"),
        "icms_aliq":      rec.get("icms_aliq"),
        "icms_valor":     rec.get("icms_valor"),
        "pis_aliq":       rec.get("pis_aliq"),
        "pis_valor":      rec.get("pis_valor"),
        "cofins_aliq":    rec.get("cofins_aliq"),
        "cofins_valor":   rec.get("cofins_valor"),
        "__triagem__":       triagem,
        "__motivos__":       " | ".join(motivos) if motivos else "",
        "__dif_tusd__":      None,
        "__dif_te__":        None,
        "__dif_band__":      None,
        "__dif_leit__":      None,
        "__dif_icms__":      None,
        "__dif_total__":     metricas.get("dif_total_R$"),
        "__dif_total_pct__": None,
        # ── Campos MT — leituras por posto ────────────────────────────────────
        "consumo_ponta_kwh":    rec.get("consumo_ponta_kwh"),
        "consumo_fp_kwh":       rec.get("consumo_fp_kwh"),
        "demanda_contratada_kw": rec.get("demanda_contratada_kw"),
        "demanda_medida_kw":    rec.get("demanda_medida_kw"),
        "demanda_ultrap_kw":    rec.get("demanda_ultrap_kw"),
        "taxa_perda":           "Sim" if rec.get("taxa_perda") else "Nao",
        # kWh Ponta
        "med_kwh_ponta_lant":   rec.get("med_kwh_ponta_lant"),
        "med_kwh_ponta_latu":   rec.get("med_kwh_ponta_latu"),
        "med_kwh_ponta_mult":   rec.get("med_kwh_ponta_mult"),
        "med_kwh_ponta_cons":   rec.get("med_kwh_ponta_cons"),
        # kWh Fora Ponta
        "med_kwh_fp_lant":      rec.get("med_kwh_fp_lant"),
        "med_kwh_fp_latu":      rec.get("med_kwh_fp_latu"),
        "med_kwh_fp_mult":      rec.get("med_kwh_fp_mult"),
        "med_kwh_fp_cons":      rec.get("med_kwh_fp_cons"),
        # kW Ponta
        "med_kw_ponta_lant":    rec.get("med_kw_ponta_lant"),
        "med_kw_ponta_latu":    rec.get("med_kw_ponta_latu"),
        "med_kw_ponta_mult":    rec.get("med_kw_ponta_mult"),
        "med_kw_ponta_cons":    rec.get("med_kw_ponta_cons"),
        # kW Fora Ponta
        "med_kw_fp_lant":       rec.get("med_kw_fp_lant"),
        "med_kw_fp_latu":       rec.get("med_kw_fp_latu"),
        "med_kw_fp_mult":       rec.get("med_kw_fp_mult"),
        "med_kw_fp_cons":       rec.get("med_kw_fp_cons"),
        # kVarh Ponta
        "med_kvarh_ponta_lant": rec.get("med_kvarh_ponta_lant"),
        "med_kvarh_ponta_latu": rec.get("med_kvarh_ponta_latu"),
        "med_kvarh_ponta_mult": rec.get("med_kvarh_ponta_mult"),
        "med_kvarh_ponta_cons": rec.get("med_kvarh_ponta_cons"),
        # kVarh Fora Ponta
        "med_kvarh_fp_lant":    rec.get("med_kvarh_fp_lant"),
        "med_kvarh_fp_latu":    rec.get("med_kvarh_fp_latu"),
        "med_kvarh_fp_mult":    rec.get("med_kvarh_fp_mult"),
        "med_kvarh_fp_cons":    rec.get("med_kvarh_fp_cons"),
    }


def _triagem_cpfl(alertas):
    cats_criticas = {"Tarifa TUSD", "Tarifa TE", "Bandeira"}
    flags_div, flags_inv = [], []
    for a in alertas:
        st  = a.get("status", "OK")
        cat = a.get("cat", "")
        if st == "INVESTIGAR":
            if cat in cats_criticas:
                dif = abs(a.get("diferenca") or 0)
                if dif >= 1.0:
                    flags_div.append("[" + cat + "] " + a.get("descricao", ""))
                else:
                    flags_inv.append("[" + cat + "] " + a.get("descricao", ""))
            else:
                flags_inv.append("[" + cat + "] " + a.get("descricao", ""))
        elif st in ("ATENCAO", "ATENÇÃO"):
            flags_inv.append("[" + cat + "] " + a.get("descricao", ""))
    if flags_div:
        return "DIVERGENCIA", flags_div + flags_inv
    if flags_inv:
        return "INVESTIGAR", flags_inv
    return "OK", []


def normalizar_cpfl(dados, audit_result, pdf_filename):
    alertas  = audit_result.get("alertas", [])
    auditado = audit_result.get("auditado", {})
    itens    = dados.get("itens", [])
    trib     = dados.get("tributos", {})
    triagem, motivos = _triagem_cpfl(alertas)
    tusd_item = next((i for i in itens if i.get("tipo") == "consumo_tusd"), {})
    te_item   = next((i for i in itens if i.get("tipo") == "consumo_te"), {})
    band_val  = sum(i.get("valor") or 0 for i in itens if i.get("tipo") == "bandeira")
    medidor   = (dados.get("medidores") or [{}])[0]

    def _dif_cat(cat):
        a = next((x for x in alertas if x.get("cat") == cat), None)
        return a.get("diferenca") if a else None

    consumo      = dados.get("consumo_faturado") or 0
    dif_tusd_tar = _dif_cat("Tarifa TUSD")
    dif_te_tar   = _dif_cat("Tarifa TE")
    dif_tusd_rs  = round(dif_tusd_tar * consumo, 2) if dif_tusd_tar and consumo else None
    dif_te_rs    = round(dif_te_tar * consumo, 2) if dif_te_tar and consumo else None
    dif_band_rs  = _dif_cat("Bandeira")
    dif_leit     = next((a.get("diferenca") for a in alertas if a.get("cat") == "Consumo Medidor"), None)
    dif_total_rs = _dif_cat("Total a Pagar")
    total_fat    = dados.get("total_fatura") or 0
    dif_total_pct = round(abs(dif_total_rs) / total_fat * 100, 1) if dif_total_rs and total_fat else None

    return {
        "arquivo":        pdf_filename,
        "distribuidora":  "CPFL Piratininga",
        "layout":         dados.get("_formato"),
        "ref_mes_ano":    dados.get("mes_ref"),
        "vencimento":     _fmt_date(dados.get("data_vencimento")),
        "conta_uc":       dados.get("conta_contrato") or dados.get("uc"),
        "cliente_nome":   dados.get("cliente_nome"),
        "subgrupo":       dados.get("subgrupo"),
        "data_emissao":   _fmt_date(dados.get("data_emissao")),
        "nr_nota_fiscal": dados.get("nota_fiscal"),
        "nr_medidor":     medidor.get("numero") or dados.get("uc"),
        "consumo_kwh":    consumo,
        "nr_dias":        dados.get("dias_ciclo"),
        "leitura_anterior": medidor.get("leitura_anterior"),
        "leitura_atual":    medidor.get("leitura_atual"),
        "preco_tusd":     tusd_item.get("preco_unit_com_trib"),
        "valor_tusd":     tusd_item.get("valor"),
        "preco_te":       te_item.get("preco_unit_com_trib"),
        "valor_te":       te_item.get("valor"),
        "tarifa_tusd_sem": auditado.get("tusd_sem_trib"),
        "tarifa_te_sem":   auditado.get("te_sem_trib"),
        "bandeira":       dados.get("bandeira_vigente"),
        "valor_bandeira": band_val or None,
        "cosip":          dados.get("cosip"),
        "total_fatura":   total_fat,
        "icms_base":      trib.get("icms", {}).get("base"),
        "icms_aliq":      trib.get("icms", {}).get("aliquota_pct"),
        "icms_valor":     trib.get("icms", {}).get("valor"),
        "pis_aliq":       trib.get("pis", {}).get("aliquota_pct"),
        "pis_valor":      trib.get("pis", {}).get("valor"),
        "cofins_aliq":    trib.get("cofins", {}).get("aliquota_pct"),
        "cofins_valor":   trib.get("cofins", {}).get("valor"),
        "__triagem__":       triagem,
        "__motivos__":       " | ".join(motivos) if motivos else "",
        "__dif_tusd__":      dif_tusd_rs,
        "__dif_te__":        dif_te_rs,
        "__dif_band__":      dif_band_rs,
        "__dif_leit__":      dif_leit,
        "__dif_icms__":      None,
        "__dif_total__":     dif_total_rs,
        "__dif_total_pct__": dif_total_pct,
    }


APP_DIR = Path(__file__).parent
ROOT = APP_DIR.parent
INPUT_DIR = ROOT / "input"
OUTPUT_DIR = ROOT / "output"
INPUT_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Minuto Energia | Auditor de Faturas",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Identidade visual Minuto Energia ──────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600&family=Inter:wght@400;500&display=swap');

/* Base */
html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
    background-color: #FAFBF8;
}
h1, h2, h3, .stTabs [data-baseweb="tab"] {
    font-family: 'Poppins', sans-serif;
}

/* Header principal */
.me-header {
    background: linear-gradient(135deg, #0A2540 0%, #1B5179 55%, #2A7456 100%);
    padding: 28px 36px;
    border-radius: 12px;
    margin-bottom: 28px;
}
.me-header h1 {
    color: #FFFFFF;
    font-size: 26px;
    font-weight: 600;
    margin: 0 0 4px 0;
    font-family: 'Poppins', sans-serif;
}
.me-header p {
    color: #C8E04A;
    font-size: 13px;
    margin: 0;
    font-family: 'Inter', sans-serif;
}

/* Métricas customizadas */
.me-metric {
    background: #FFFFFF;
    border-radius: 10px;
    padding: 18px 20px;
    border-left: 4px solid #1B4173;
    box-shadow: 0 1px 4px rgba(10,37,64,0.08);
}
.me-metric.ok    { border-left-color: #5A9F37; }
.me-metric.inv   { border-left-color: #1B5179; }
.me-metric.div   { border-left-color: #BA7517; }
.me-metric .val  {
    font-family: 'Inter', sans-serif;
    font-size: 28px;
    font-weight: 500;
    color: #0A2540;
    font-feature-settings: "tnum";
}
.me-metric .lbl  {
    font-size: 12px;
    color: #5A6B7C;
    margin-top: 2px;
}

/* Badges de triagem */
.badge {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 20px;
    font-size: 11px;
    font-weight: 500;
    font-family: 'Inter', sans-serif;
}
.badge-ok   { background: #F5F9F2; color: #5A9F37; border: 1px solid #5A9F37; }
.badge-inv  { background: #EEF3FB; color: #1B5179; border: 1px solid #1B5179; }
.badge-div  { background: #FDF6EB; color: #BA7517; border: 1px solid #BA7517; }

/* Drop-zone */
[data-testid="stFileUploader"] {
    border: 2px dashed #5A9F37 !important;
    border-radius: 10px;
    background: #FAFBF8;
    padding: 8px;
}

/* Botão primário */
.stButton > button[kind="primary"] {
    background-color: #1B4173;
    color: #FFFFFF;
    border: none;
    border-radius: 8px;
    font-family: 'Poppins', sans-serif;
    font-weight: 500;
    padding: 10px 24px;
}
.stButton > button[kind="primary"]:hover {
    background-color: #0A2540;
}

/* Botão secundário */
.stButton > button[kind="secondary"] {
    border: 1.5px solid #1B4173;
    color: #1B4173;
    border-radius: 8px;
    font-family: 'Poppins', sans-serif;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background-color: #0A2540;
}
[data-testid="stSidebar"] * {
    color: #FFFFFF !important;
}
[data-testid="stSidebar"] .stMarkdown p {
    color: #C5D7E5 !important;
}

/* Tabs */
.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
    border-bottom: 2px solid #E5EBE0;
}
.stTabs [data-baseweb="tab"] {
    padding: 8px 20px;
    color: #5A6B7C;
    font-size: 14px;
}
.stTabs [aria-selected="true"] {
    color: #1B4173 !important;
    border-bottom: 2px solid #1B4173 !important;
    font-weight: 500;
}

/* Progress bar */
.stProgress > div > div {
    background-color: #5A9F37;
}

/* Dataframe */
[data-testid="stDataFrame"] {
    border: 1px solid #E5EBE0;
    border-radius: 8px;
}

/* Divider */
hr { border-color: #E5EBE0; }

/* Download button */
.stDownloadButton > button {
    background-color: #5A9F37;
    color: white;
    border: none;
    border-radius: 8px;
    font-family: 'Poppins', sans-serif;
    font-weight: 500;
}
.stDownloadButton > button:hover {
    background-color: #4a8a2e;
}

/* Botão Limpar — fundo escuro, texto amarelo-verde */
button[data-testid="stBaseButton-secondary"] {
    background-color: #0A2540 !important;
    color: #C8E04A !important;
    border: none !important;
    font-weight: 700 !important;
}
button[data-testid="stBaseButton-secondary"]:hover {
    background-color: #0d2f50 !important;
    color: #C8E04A !important;
    opacity: 0.9;
}
</style>
""", unsafe_allow_html=True)


# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="me-header">
    <h1>⚡ Minuto Energia — Auditor de Faturas</h1>
    <p>Auditoria automática conforme REH ANEEL e Lei 14.300/2022</p>
</div>
""", unsafe_allow_html=True)


# ── COSIP check ───────────────────────────────────────────────────────────────
def aplicar_check_cosip(registros, tolerancia=0.10):
    from collections import defaultdict
    uc_cosips = defaultdict(list)
    for reg in registros:
        uc = reg.get("conta_uc")
        cosip = reg.get("cosip")
        if uc and cosip and cosip > 0:
            uc_cosips[uc].append(cosip)
    for reg in registros:
        uc = reg.get("conta_uc")
        cosip = reg.get("cosip")
        if not uc or not cosip or cosip <= 0:
            continue
        valores = uc_cosips.get(uc, [])
        if len(valores) < 2:
            continue
        media = sum(valores) / len(valores)
        limite = round(media * (1 + tolerancia), 2)
        if cosip > limite:
            if reg.get("__triagem__") != "DIVERGENCIA":
                reg["__triagem__"] = "INVESTIGAR"
            motivo = f"COSIP R${cosip:.2f} > +10% da media historica UC (media R${media:.2f}, limite R${limite:.2f})"
            atual = reg.get("__motivos__", "")
            reg["__motivos__"] = (atual + " | " + motivo).lstrip(" | ") if atual else motivo
    return registros


# ── Sidebar ───────────────────────────────────────────────────────────────────
def safe_filename(s):
    s = re.sub(r"[^\w\d\-_]", "_", str(s))
    return s[:80]


def processar_fatura_neo(pdf_path, run_dir):
    pdf_path = Path(pdf_path)
    r = neo_parse(str(pdf_path))
    triagem, motivos, metricas = neo_auditar(r)

    # Registro normalizado para Excel-mestre
    r_norm = dict(r)
    r_norm["arquivo"] = pdf_path.name
    registro = normalizar_neoenergia_pe(r_norm, triagem, motivos, metricas)

    # Formato compatível com UI (alertas por item)
    st_map = {"OK": "OK", "INVESTIGAR": "INVESTIGAR", "DIVERGENCIA": "ATENÇÃO"}
    st = st_map.get(triagem, "OK")
    alertas = (
        [{"cat": "Auditoria", "descricao": m, "status": st} for m in motivos]
        if motivos
        else [{"cat": "Auditoria", "descricao": "Fatura conferida", "status": "OK"}]
    )

    dados = {
        "cliente_nome":     r.get("conta_contrato", ""),
        "uc":               r.get("conta_contrato", ""),
        "subgrupo":         "B3",
        "leitura_anterior": None,
        "leitura_atual":    None,
        "dias_ciclo":       r.get("nr_dias"),
        "total_fatura":     r.get("total_fatura"),
        "total_a_pagar":    r.get("total_fatura"),
        "mes_ref":          r.get("ref_mes_ano", ""),
        "nota_fiscal":      r.get("nr_nota_fiscal", ""),
    }

    return {
        "pdf_filename":        pdf_path.name,
        "dados":               dados,
        "config":              {},
        "audit_result":        {"alertas": alertas, "auditado": {"reh_aplicada": ""}},
        "excel_individual_path": "",
        "__registro__":        registro,
    }


def processar_fatura_paulista(pdf_path, run_dir):
    pdf_path = Path(pdf_path)
    r = paulista_parse(str(pdf_path))
    triagem, motivos, metricas = paulista_auditar(r)

    registro = normalizar_cpfl_paulista(r, triagem, motivos, metricas)

    st_map = {"OK": "OK", "INVESTIGAR": "INVESTIGAR", "DIVERGENCIA": "ATENÇÃO"}
    st = st_map.get(triagem, "OK")
    alertas = (
        [{"cat": "Auditoria", "descricao": m, "status": st} for m in motivos]
        if motivos
        else [{"cat": "Auditoria", "descricao": "Fatura conferida", "status": "OK"}]
    )

    dados = {
        "cliente_nome":     r.get("cliente_nome", ""),
        "uc":               r.get("conta_contrato", ""),
        "subgrupo":         r.get("subgrupo", "A4"),
        "leitura_anterior": None,
        "leitura_atual":    None,
        "dias_ciclo":       r.get("nr_dias"),
        "total_fatura":     r.get("total_fatura"),
        "total_a_pagar":    r.get("total_a_pagar"),
        "mes_ref":          r.get("ref_mes_ano", ""),
        "nota_fiscal":      r.get("nr_nota_fiscal", ""),
    }

    return {
        "pdf_filename":          pdf_path.name,
        "dados":                 dados,
        "config":                {},
        "audit_result":          {"alertas": alertas, "auditado": {"reh_aplicada": metricas.get("reh_aplicado", "")}},
        "excel_individual_path": "",
        "__registro__":          registro,
    }


def processar_fatura(pdf_path, config_default, run_dir):
    pdf_path = Path(pdf_path)
    dados = extract_fatura(str(pdf_path))

    config = dict(config_default)
    config["tem_gd"] = dados.get("tem_gd", False)
    if config.get("usar_cat_como_compensada", True):
        config["energia_compensada_kwh"] = dados.get("gd_ajuste_cat") or dados.get("gd_injetada_mes")

    audit_result = auditar_fatura(dados, config)

    uc = safe_filename(dados.get("uc", "UC"))
    mes = (dados.get("mes_ref", "") or "").replace("/", "-")
    nf = safe_filename(dados.get("nota_fiscal", "NF"))
    out_xlsx = run_dir / f"Auditoria_{uc}_{mes}_NF{nf}.xlsx"
    preencher_template(dados, config, audit_result, out_xlsx)

    registro = normalizar_cpfl(dados, audit_result, pdf_path.name)

    return {
        "pdf_filename":        pdf_path.name,
        "dados":               dados,
        "config":              config,
        "audit_result":        audit_result,
        "excel_individual_path": str(out_xlsx),
        "__registro__":        registro,
    }


with st.sidebar:
    # ── Nome do cliente (topo da sidebar) ────────────────────────────────────
    st.markdown("### Cliente")
    cliente_nome = st.text_input(
        "Nome do cliente",
        value=st.session_state.get("cliente_nome", ""),
        placeholder="Digite o nome do cliente",
        label_visibility="collapsed",
    )
    st.session_state["cliente_nome"] = cliente_nome

    parceiro_nome = st.text_input(
        "Nome do parceiro (opcional)",
        value=st.session_state.get("parceiro_nome", ""),
        placeholder="Deixe em branco para relatório Minuto apenas",
        label_visibility="visible",
    )
    st.session_state["parceiro_nome"] = parceiro_nome

    # Contador de registros acumulados
    n_acum = len(st.session_state.get("registros_acumulados", []))
    if n_acum > 0:
        n_ucs_acum = len({r.get("conta_uc") for r in st.session_state["registros_acumulados"]})
        st.caption(f"📂 {n_acum} faturas acumuladas · {n_ucs_acum} UC(s)")
        if st.button("🗑️ Nova análise", key="nova_analise"):
            st.session_state["registros_acumulados"] = []
            st.session_state["cliente_nome"] = ""
            st.session_state["parceiro_nome"] = ""
            st.session_state.pop("last_run", None)
            st.session_state["_uploader_key"] = st.session_state.get("_uploader_key", 0) + 1
            st.rerun()

    st.markdown("---")
    st.markdown("### Configuração do Lote")
    st.markdown("---")

    opcoes_dist = ["CPFL Piratininga", "CPFL Paulista", "Neoenergia PE"]
    distribuidora = st.selectbox("Distribuidora", opcoes_dist)

    # Auto-limpar ao trocar distribuidora
    if st.session_state.get("_dist_anterior") != distribuidora:
        st.session_state["_dist_anterior"] = distribuidora
        st.session_state.pop("last_run", None)
        st.session_state["_uploader_key"] = st.session_state.get("_uploader_key", 0) + 1

    st.markdown("---")
    data_adesao = st.date_input(
        "Data de adesão MMGD",
        value=datetime.date(2022, 1, 1),
        help="Sistemas com adesão até 06/01/2023 são pré-MMGD (isentos de Fio B até 2045 — Art. 26 Lei 14.300/2022)"
    )
    usar_cat = st.checkbox(
        "Usar ajuste CAT como energia compensada",
        value=True,
    )
    st.markdown("---")
    st.markdown('<p style="font-size:11px; color:#9AA8B7;">REH e bandeiras: pasta app/data/</p>', unsafe_allow_html=True)

config_default = {
    "data_adesao_mmgd": data_adesao.strftime("%Y-%m-%d"),
    "usar_cat_como_compensada": usar_cat,
}


# ── Tabs ──────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs(["📤 Processar faturas", "📊 Resultados", "📋 Relatório Final", "📈 Histórico", "ℹ️ Sobre"])


# ── Tab 1: Upload ─────────────────────────────────────────────────────────────
with tab1:
    st.markdown("#### Selecione as faturas em PDF")
    st.caption("Formatos aceitos: faturas em PDF")

    _ukey = st.session_state.get("_uploader_key", 0)
    uploaded = st.file_uploader(
        "Arraste os PDFs aqui ou clique para selecionar",
        type=["pdf"],
        accept_multiple_files=True,
        label_visibility="collapsed",
        key=f"uploader_{_ukey}",
    )

    if uploaded:
        st.markdown(f"**{len(uploaded)} fatura(s) selecionada(s)**")
        cols = st.columns(3)
        for i, u in enumerate(uploaded):
            cols[i % 3].markdown(f"📄 {u.name} · {u.size // 1024} KB")

        st.markdown("")
        col_btn1, col_btn2 = st.columns([3, 1])
        processar = col_btn1.button(f"Processar {len(uploaded)} fatura(s)", type="primary")
        if col_btn2.button("🗑️ Limpar"):
            st.session_state.pop("last_run", None)
            st.session_state["_uploader_key"] = _ukey + 1
            st.rerun()
        if processar:
            run_id = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            run_dir = OUTPUT_DIR / f"run_{run_id}"
            run_dir.mkdir(parents=True, exist_ok=True)

            progress = st.progress(0)
            status = st.empty()
            resultados = []
            erros = []

            for i, file in enumerate(uploaded):
                status.markdown(f"🔍 Analisando **{file.name}** ({i+1}/{len(uploaded)})...")
                tmp_pdf = run_dir / file.name
                with open(tmp_pdf, "wb") as f:
                    f.write(file.read())
                try:
                    if distribuidora == "Neoenergia PE":
                        res = processar_fatura_neo(tmp_pdf, run_dir)
                    elif distribuidora == "CPFL Paulista":
                        res = processar_fatura_paulista(tmp_pdf, run_dir)
                    else:
                        res = processar_fatura(tmp_pdf, config_default, run_dir)
                    resultados.append(res)
                except Exception as e:
                    erros.append({"file": file.name, "error": str(e)})
                progress.progress((i + 1) / len(uploaded))

            status.markdown("📊 Consolidando resultados...")
            master_path = run_dir / f"MASTER_Auditoria_{run_id}.xlsx"
            registros_norm = [r["__registro__"] for r in resultados]
            registros_norm = aplicar_check_cosip(registros_norm)
            gerar_excel_mestre(registros_norm, master_path)

            status.markdown(f"✅ **Auditoria conferida!** {len(resultados)} fatura(s) processada(s).")

            # ── Acumular registros em session_state (deduplica por arquivo) ──
            acumulados = st.session_state.get("registros_acumulados", [])
            arquivos_existentes = {r.get("arquivo") for r in acumulados}
            novos = [r["__registro__"] for r in resultados
                     if r["__registro__"].get("arquivo") not in arquivos_existentes]
            st.session_state["registros_acumulados"] = acumulados + novos

            st.session_state["last_run"] = {
                "run_id": run_id,
                "run_dir": str(run_dir),
                "resultados": resultados,
                "erros": erros,
                "master_path": str(master_path),
            }

            if erros:
                st.warning(f"Atenção necessária em {len(erros)} arquivo(s):")
                for e in erros:
                    st.markdown(f"- **{e['file']}**: {e['error']}")

            st.info("Acesse a aba **Resultados** para visualizar e baixar os relatórios.")


# ── Tab 2: Resultados ─────────────────────────────────────────────────────────
with tab2:
    if "last_run" not in st.session_state:
        st.markdown("""
        <div style="text-align:center; padding: 60px 0; color: #5A6B7C;">
            <div style="font-size:48px; margin-bottom:16px;">📂</div>
            <div style="font-family:'Poppins',sans-serif; font-size:16px; font-weight:500; color:#0A2540;">
                Nenhum lote processado ainda
            </div>
            <div style="font-size:13px; margin-top:8px;">
                Processe as faturas na aba <strong>Processar faturas</strong>
            </div>
        </div>
        """, unsafe_allow_html=True)
    else:
        run = st.session_state["last_run"]
        resultados = run["resultados"]

        total_ok   = sum(1 for r in resultados if r["__registro__"]["__triagem__"] == "OK")
        total_inv  = sum(1 for r in resultados if r["__registro__"]["__triagem__"] == "INVESTIGAR")
        total_div  = sum(1 for r in resultados if r["__registro__"]["__triagem__"] == "DIVERGENCIA")
        total_valor = sum(r["__registro__"].get("total_fatura") or 0 for r in resultados)
        valor_fmt = f"R$ {total_valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        # Métricas
        c1, c2, c3, c4 = st.columns(4)
        c1.markdown(f"""
        <div class="me-metric">
            <div class="val">{len(resultados)}</div>
            <div class="lbl">Faturas processadas</div>
        </div>""", unsafe_allow_html=True)
        c2.markdown(f"""
        <div class="me-metric ok">
            <div class="val">{total_ok}</div>
            <div class="lbl">Conferidos — OK</div>
        </div>""", unsafe_allow_html=True)
        c3.markdown(f"""
        <div class="me-metric inv">
            <div class="val">{total_inv}</div>
            <div class="lbl">Investigar</div>
        </div>""", unsafe_allow_html=True)
        c4.markdown(f"""
        <div class="me-metric div">
            <div class="val">{total_div}</div>
            <div class="lbl">Divergência</div>
        </div>""", unsafe_allow_html=True)

        st.markdown(f"""
        <div style="margin:20px 0 8px; font-family:'Inter',sans-serif;">
            <span style="font-size:13px; color:#5A6B7C;">Valor total do lote</span><br>
            <span style="font-size:28px; font-weight:500; color:#0A2540; font-feature-settings:'tnum';">{valor_fmt}</span>
        </div>
        """, unsafe_allow_html=True)

        st.divider()

        # Tabela resumo
        rows = []
        for r in resultados:
            reg = r["__registro__"]
            rows.append({
                "Arquivo":     reg.get("arquivo", r["pdf_filename"]),
                "Distribuidora": reg.get("distribuidora", ""),
                "Cliente":     (reg.get("cliente_nome") or reg.get("conta_uc") or "")[:30],
                "UC":          reg.get("conta_uc", ""),
                "Ref Mes/Ano": reg.get("ref_mes_ano", ""),
                "Dias":        reg.get("nr_dias"),
                "Total (R$)":  reg.get("total_fatura"),
                "Triagem":     reg.get("__triagem__", ""),
            })
        df = pd.DataFrame(rows)
        st.dataframe(df, use_container_width=True, hide_index=True)

        st.divider()

        # Downloads
        st.markdown("#### Baixar relatórios")
        col1, col2 = st.columns(2)

        with col1:
            with open(run["master_path"], "rb") as f:
                st.download_button(
                    "📊 Baixar Excel-mestre",
                    f.read(),
                    file_name=f"Auditoria_Minuto_Energia_{run['run_id']}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        with col2:
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.write(run["master_path"], Path(run["master_path"]).name)
                for r in resultados:
                    p = r["excel_individual_path"]
                    if p:
                        zf.write(p, Path(p).name)
            buf.seek(0)
            st.download_button(
                "📦 Baixar ZIP completo",
                buf.getvalue(),
                file_name=f"Auditoria_Completa_{run['run_id']}.zip",
                mime="application/zip",
            )

        st.divider()

        # Alertas por fatura
        st.markdown("#### Detalhamento por fatura")
        for r in resultados:
            reg    = r["__registro__"]
            triagem = reg.get("__triagem__", "OK")
            motivos_str = reg.get("__motivos__", "")
            uc_label = reg.get("conta_uc") or r["dados"].get("uc", "")
            cliente_label = (reg.get("cliente_nome") or r["dados"].get("cliente_nome") or "")[:30]
            label = f"{r['pdf_filename']} · {cliente_label} · UC {uc_label}"
            badge = {"INVESTIGAR": " 🔍", "DIVERGENCIA": " ⚠️"}.get(triagem, " ✅")
            with st.expander(label + badge):
                icon = {"OK": "✅", "INVESTIGAR": "🔍", "DIVERGENCIA": "⚠️"}.get(triagem, "❓")
                if motivos_str:
                    for m in motivos_str.split(" | "):
                        st.markdown(f"{icon} {m}")
                else:
                    st.markdown("✅ Fatura conferida — sem divergências")


# ── Tab 3: Relatório Final ────────────────────────────────────────────────────
with tab3:
    registros_acum = st.session_state.get("registros_acumulados", [])

    if not registros_acum:
        st.markdown("""
        <div style="text-align:center; padding: 60px 0; color: #5A6B7C;">
            <div style="font-size:48px; margin-bottom:16px;">📋</div>
            <div style="font-family:'Poppins',sans-serif; font-size:16px; font-weight:500; color:#0A2540;">
                Nenhuma fatura acumulada ainda
            </div>
            <div style="font-size:13px; margin-top:8px;">
                Processe lotes na aba <strong>Processar faturas</strong> e volte aqui quando concluir o envio.
            </div>
        </div>
        """, unsafe_allow_html=True)
    else:
        n_fat   = len(registros_acum)
        n_ucs   = len({r.get("conta_uc") for r in registros_acum})
        n_ok    = sum(1 for r in registros_acum if r.get("__triagem__") == "OK")
        n_inv   = sum(1 for r in registros_acum if r.get("__triagem__") == "INVESTIGAR")
        n_div   = sum(1 for r in registros_acum if r.get("__triagem__") == "DIVERGENCIA")
        val_tot = sum(r.get("total_fatura") or 0 for r in registros_acum)
        val_fmt = f"R$ {val_tot:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        nome_exibido = st.session_state.get("cliente_nome", "").strip() or "*(nome não informado)*"
        st.markdown(f"**Cliente:** {nome_exibido}")
        st.markdown("")

        c1, c2, c3, c4 = st.columns(4)
        c1.markdown(f'<div class="me-metric"><div class="val">{n_fat}</div><div class="lbl">Faturas acumuladas</div></div>', unsafe_allow_html=True)
        c2.markdown(f'<div class="me-metric ok"><div class="val">{n_ok}</div><div class="lbl">OK</div></div>', unsafe_allow_html=True)
        c3.markdown(f'<div class="me-metric inv"><div class="val">{n_inv}</div><div class="lbl">Investigar</div></div>', unsafe_allow_html=True)
        c4.markdown(f'<div class="me-metric div"><div class="val">{n_div}</div><div class="lbl">Divergência</div></div>', unsafe_allow_html=True)

        st.markdown(f"""
        <div style="margin:16px 0 8px;">
            <span style="font-size:13px; color:#5A6B7C;">Valor total acumulado</span><br>
            <span style="font-size:28px; font-weight:500; color:#0A2540; font-feature-settings:'tnum';">{val_fmt}</span>
        </div>
        """, unsafe_allow_html=True)

        st.markdown(f"**{n_ucs} unidade(s) consumidora(s)**")
        ucs_labels = sorted({str(r.get("conta_uc") or "—") for r in registros_acum})
        st.caption("  ·  ".join(ucs_labels))

        st.divider()

        # Validação antes de gerar
        cliente_val = st.session_state.get("cliente_nome", "").strip()
        if not cliente_val:
            st.warning("⚠️ Informe o **nome do cliente** na barra lateral antes de gerar o relatório.")

        col_btn, col_info = st.columns([2, 3])
        with col_btn:
            gerar = st.button(
                "📄 Envio finalizado — Gerar Relatório PDF",
                type="primary",
                disabled=(not cliente_val or not RELATORIO_PDF_DISPONIVEL),
            )

        if gerar and cliente_val:
            with st.spinner("Gerando relatório PDF..."):
                try:
                    parceiro_val = st.session_state.get("parceiro_nome", "").strip()
                    pdf_bytes = gerar_relatorio_pdf(cliente_val, registros_acum, parceiro_nome=parceiro_val)
                    nome_arquivo = re.sub(r"[^\w\d\-_]", "_", cliente_val)[:40]
                    st.success("✅ Relatório gerado com sucesso!")
                    st.download_button(
                        "⬇️ Baixar Relatório PDF",
                        pdf_bytes,
                        file_name=f"Relatorio_Auditoria_{nome_arquivo}.pdf",
                        mime="application/pdf",
                    )
                except Exception as e:
                    st.error(f"Erro ao gerar relatório: {e}")


# ── Tab 4: Histórico de Consumo ───────────────────────────────────────────────
with tab4:
    try:
        import plotly.graph_objects as go
        _PLOTLY_OK = True
    except ImportError:
        _PLOTLY_OK = False

    _registros_hist = st.session_state.get("registros_acumulados", [])
    _df_hist_all = pd.DataFrame([
        {
            "uc":           r.get("conta_uc") or "—",
            "distribuidora": r.get("distribuidora", ""),
            "ref":          r.get("ref_mes_ano") or "",
            "consumo_kwh":  r.get("consumo_kwh"),
            "total_fatura": r.get("total_fatura"),
            "arquivo":      r.get("arquivo", ""),
        }
        for r in _registros_hist
        if r.get("consumo_kwh") is not None
    ])

    if _df_hist_all.empty:
        st.markdown("""
        <div style="text-align:center; padding: 60px 0; color: #5A6B7C;">
            <div style="font-size:48px; margin-bottom:16px;">📈</div>
            <div style="font-family:'Poppins',sans-serif; font-size:16px; font-weight:500; color:#0A2540;">
                Nenhum dado de consumo disponível
            </div>
            <div style="font-size:13px; margin-top:8px;">
                Processe faturas na aba <strong>Processar faturas</strong> para ver o histórico.
            </div>
        </div>
        """, unsafe_allow_html=True)
    else:
        def _parse_ref(ref):
            try:
                m, y = ref.strip().split("/")
                return pd.Timestamp(year=int(y), month=int(m), day=1)
            except Exception:
                return pd.NaT

        _df_hist_all["data"] = _df_hist_all["ref"].apply(_parse_ref)
        _df_hist_all = _df_hist_all.sort_values("data")

        _ucs = sorted(_df_hist_all["uc"].unique())

        # Resumo por UC (quando há mais de uma)
        if len(_ucs) > 1:
            st.markdown("#### Resumo por unidade consumidora")
            _resumo = []
            for _uc in _ucs:
                _dfu = _df_hist_all[_df_hist_all["uc"] == _uc]
                _med = _dfu["consumo_kwh"].mean()
                _na  = int((((_dfu["consumo_kwh"] - _med).abs() / _med) > 0.15).sum()) if _med else 0
                _resumo.append({
                    "UC": _uc,
                    "Distribuidora": _dfu["distribuidora"].iloc[0],
                    "Meses": len(_dfu),
                    "Média (kWh)": round(_med, 1),
                    "Anomalias": _na,
                })
            st.dataframe(pd.DataFrame(_resumo), use_container_width=True, hide_index=True)
            st.divider()

        _uc_sel = st.selectbox("Selecionar Unidade Consumidora", _ucs) if len(_ucs) > 1 else _ucs[0]
        _df_uc  = _df_hist_all[_df_hist_all["uc"] == _uc_sel].copy()

        _media = _df_uc["consumo_kwh"].mean()
        _df_uc["desvio_pct"] = ((_df_uc["consumo_kwh"] - _media) / _media * 100).round(1)
        _df_uc["anomalia"]   = _df_uc["consumo_kwh"].apply(
            lambda x: abs(x - _media) / _media > 0.15 if _media else False
        )

        _n_meses = len(_df_uc)
        _n_anom  = int(_df_uc["anomalia"].sum())

        # Métricas
        _ca, _cb, _cc = st.columns(3)
        _ca.markdown(f'<div class="me-metric"><div class="val">{_n_meses}</div><div class="lbl">Meses no histórico</div></div>', unsafe_allow_html=True)
        _cb.markdown(f'<div class="me-metric ok"><div class="val">{_media:.0f} kWh</div><div class="lbl">Consumo médio</div></div>', unsafe_allow_html=True)
        _cc.markdown(f'<div class="me-metric {"div" if _n_anom else "ok"}"><div class="val">{_n_anom}</div><div class="lbl">Meses com anomalia (±15%)</div></div>', unsafe_allow_html=True)

        st.markdown("")

        # Gráfico
        if _PLOTLY_OK:
            _df_ok   = _df_uc[~_df_uc["anomalia"]]
            _df_anom = _df_uc[_df_uc["anomalia"]]

            _fig = go.Figure()

            # Linha de consumo
            _fig.add_trace(go.Scatter(
                x=_df_uc["ref"], y=_df_uc["consumo_kwh"],
                mode="lines",
                line=dict(color="#1B5179", width=2),
                showlegend=False, hoverinfo="skip",
            ))

            # Pontos normais
            _fig.add_trace(go.Scatter(
                x=_df_ok["ref"], y=_df_ok["consumo_kwh"],
                mode="markers",
                marker=dict(color="#5A9F37", size=9),
                name="Normal",
                hovertemplate="%{x}<br><b>%{y:.0f} kWh</b><extra></extra>",
            ))

            # Pontos anômalos
            if not _df_anom.empty:
                _fig.add_trace(go.Scatter(
                    x=_df_anom["ref"], y=_df_anom["consumo_kwh"],
                    mode="markers",
                    marker=dict(color="#BA7517", size=13, symbol="circle",
                                line=dict(width=2, color="#7A4D0E")),
                    name="Anomalia (±15%)",
                    customdata=_df_anom["desvio_pct"],
                    hovertemplate="%{x}<br><b>%{y:.0f} kWh</b><br>Desvio: %{customdata:+.1f}%<extra></extra>",
                ))

            # Linha da média e faixas
            _fig.add_hline(y=_media, line_dash="dash", line_color="#9AA8B7",
                           annotation_text=f"Média {_media:.0f} kWh",
                           annotation_position="top right")
            _fig.add_hrect(y0=_media * 0.85, y1=_media * 1.15,
                           fillcolor="#5A9F37", opacity=0.06, line_width=0)
            _fig.add_hline(y=_media * 1.15, line_dash="dot", line_color="#BA7517", opacity=0.5,
                           annotation_text="+15%", annotation_position="top right")
            _fig.add_hline(y=_media * 0.85, line_dash="dot", line_color="#BA7517", opacity=0.5,
                           annotation_text="−15%", annotation_position="bottom right")

            _fig.update_layout(
                title=dict(text=f"Histórico de Consumo — UC {_uc_sel}",
                           font=dict(family="Poppins", size=15, color="#0A2540")),
                xaxis_title="Mês/Ano",
                yaxis_title="Consumo (kWh)",
                plot_bgcolor="#FAFBF8",
                paper_bgcolor="#FAFBF8",
                legend=dict(orientation="h", y=1.1, x=0),
                height=420,
                margin=dict(t=70, r=110),
            )
            st.plotly_chart(_fig, use_container_width=True)
        else:
            st.line_chart(_df_uc.set_index("ref")["consumo_kwh"])

        # Alertas textuais
        if _n_anom:
            for _, _row in _df_uc[_df_uc["anomalia"]].iterrows():
                _dir = "acima" if _row["desvio_pct"] > 0 else "abaixo"
                st.warning(f"⚠️ **{_row['ref']}** — {_row['consumo_kwh']:.0f} kWh · {abs(_row['desvio_pct']):.1f}% {_dir} da média")
        else:
            st.success(f"✅ Consumo estável nos {_n_meses} meses analisados")

        # Tabela completa
        st.divider()
        st.markdown("#### Histórico completo")
        _df_tab = _df_uc[["ref", "consumo_kwh", "desvio_pct", "total_fatura", "arquivo"]].copy()
        _df_tab.columns = ["Mês/Ano", "Consumo (kWh)", "Desvio (%)", "Total Fatura (R$)", "Arquivo"]
        _df_tab.insert(3, "Anomalia", _df_uc["anomalia"].map({True: "⚠️ Sim", False: "✅ Não"}).values)
        st.dataframe(_df_tab, use_container_width=True, hide_index=True)


# ── Tab 5: Sobre ──────────────────────────────────────────────────────────────
with tab5:
    st.markdown("#### Como funciona")
    st.markdown("""
A ferramenta extrai, audita e classifica cada item da fatura de forma automática:

1. Você carrega os PDFs das faturas
2. O sistema extrai todos os dados: cliente, UC, leituras, valores, bandeiras, tributos
3. Recalcula tarifas conforme a REH ANEEL vigente na data da fatura
4. Compara o cobrado vs o auditado e classifica como **Conferido / Investigar / Atenção**
5. Gera Excel-mestre consolidado + individuais detalhados por fatura
""")

    st.divider()
    st.markdown("#### Verificações automáticas")
    st.markdown("""
- Tarifa TUSD/TE com gross-up de tributos vs REH vigente
- Tese do Século — exclusão do ICMS da base PIS/COFINS (RE 574.706/STF)
- ICMS calculado sobre a base correta
- Bandeira tarifária proporcional aos dias em cada patamar
- Fio B — Lei 14.300/2022 (pre-MMGD isento, pos-MMGD escalonado)
- Periodo de leitura dentro dos limites da REN 1.000/2021 (15-45 dias)
- Cobranças retroativas (juros, multa, atualização monetária)
- Divergência entre total da fatura e total a pagar
- GD: compensação superior à energia injetada
""")

    st.divider()
    st.markdown("#### REHs cadastradas")
    st.markdown("""
| REH | Vigência | TUSD (R$/kWh) | TE (R$/kWh) |
|---|---|---|---|
| 3409/2024 | 23/10/2024 – 22/10/2025 | 0,37008 | 0,32865 |
| 3543/2025 | 23/10/2025 – 22/10/2026 | 0,39564 | 0,34405 |
""")

    st.divider()
    st.caption("Minuto Energia - Gestao e Eficiencia Energetica - minutoenergia.com.br")

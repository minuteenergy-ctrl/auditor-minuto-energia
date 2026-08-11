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
_RELATORIO_PDF_ERRO = None
try:
    from relatorio_pdf import gerar_relatorio_pdf
    RELATORIO_PDF_DISPONIVEL = True
except Exception as _e:
    RELATORIO_PDF_DISPONIVEL = False
    _RELATORIO_PDF_ERRO = str(_e)

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
        "__dif_total__":     (metricas.get("dif_total_R$") or 0) + (metricas.get("dif_scee_R$") or 0),
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
        # ── Reativo excedente ──────────────────────────────────────────────────
        "consumo_reativo_exc_ponta_kwh": rec.get("consumo_reativo_exc_ponta_kwh"),
        "consumo_reativo_exc_fp_kwh":    rec.get("consumo_reativo_exc_fp_kwh"),
        "valor_reativo_exc_ponta":       rec.get("valor_reativo_exc_ponta"),
        "valor_reativo_exc_fp":          rec.get("valor_reativo_exc_fp"),
        # ── USDG — Demanda de Geração ──────────────────────────────────────────
        "usdg_kw":            rec.get("usdg_kw"),
        "valor_usdg":         rec.get("valor_usdg"),
        "usdg_ultrap_kw":     rec.get("usdg_ultrap_kw"),
        "usdg_ultrap_sem":    rec.get("usdg_ultrap_sem"),
        "usdg_ultrap_com":    rec.get("usdg_ultrap_com"),
        "valor_usdg_ultrap":  rec.get("valor_usdg_ultrap"),
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
tab1, tab2, tab_dash, tab3, tab4, tab5 = st.tabs([
    "📤 Processar faturas",
    "📊 Resultados",
    "📈 Dashboard",
    "📋 Relatório Final",
    "🕐 Histórico",
    "ℹ️ Sobre",
])


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
            registros_norm = [r["__registro__"] for r in resultados]
            registros_norm = aplicar_check_cosip(registros_norm)

            def _is_mt(r):
                sub = str(r.get("subgrupo", "")).upper()
                return sub.startswith("A") or r.get("layout") == "Verde-A4"

            regs_mt = [r for r in registros_norm if _is_mt(r)]
            regs_bt = [r for r in registros_norm if not _is_mt(r)]
            master_paths = []
            if regs_mt:
                mt_path = run_dir / f"MASTER_MT_Auditoria_{run_id}.xlsx"
                gerar_excel_mestre(regs_mt, mt_path, modelo="MT")
                master_paths.append(str(mt_path))
            if regs_bt:
                bt_path = run_dir / f"MASTER_BT_Auditoria_{run_id}.xlsx"
                gerar_excel_mestre(regs_bt, bt_path, modelo="BT")
                master_paths.append(str(bt_path))

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
                "master_paths": master_paths,
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
        master_paths = run.get("master_paths", [])
        excel_cols = st.columns(max(len(master_paths), 1) + 1)

        for i, mp in enumerate(master_paths):
            tipo = "MT" if "MASTER_MT" in mp else "BT"
            with excel_cols[i]:
                with open(mp, "rb") as f:
                    st.download_button(
                        f"📊 Excel {tipo}",
                        f.read(),
                        file_name=Path(mp).name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

        with excel_cols[-1]:
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for mp in master_paths:
                    zf.write(mp, Path(mp).name)
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


# ── Tab Dashboard ─────────────────────────────────────────────────────────────
with tab_dash:
    try:
        import plotly.graph_objects as _go
        import plotly.express as _px
        _PLOTLY_DASH = True
    except ImportError:
        _PLOTLY_DASH = False

    _MESES_DASH = {
        "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4,
        "MAI": 5, "JUN": 6, "JUL": 7, "AGO": 8,
        "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12,
    }

    def _parse_ref_dash(ref):
        try:
            parts = str(ref).strip().split("/")
            y = int(parts[1])
            m_raw = parts[0]
            m = int(m_raw) if m_raw.isdigit() else _MESES_DASH.get(m_raw.upper(), 1)
            return pd.Timestamp(year=y, month=m, day=1)
        except Exception:
            return pd.NaT

    def _brl_d(v):
        return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def _col(df, name):
        """Retorna coluna como Series numérica ou zeros."""
        if name in df.columns:
            return pd.to_numeric(df[name], errors="coerce").fillna(0)
        return pd.Series([0] * len(df), index=df.index)

    # ── CSS dark dashboard ────────────────────────────────────────────────────
    st.markdown("""
    <style>
    /* ── Dashboard dark cards ─────────────────────────────────────────────── */
    .dash-section-title {
        font-family: 'Poppins', sans-serif;
        font-size: 12px;
        font-weight: 600;
        color: #C8E04A;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        margin: 6px 0 14px 0;
        padding-left: 10px;
        border-left: 3px solid #C8E04A;
    }
    .dash-kpi {
        background: #0A2540;
        border-radius: 12px;
        padding: 15px 18px 13px;
        border: 1px solid #1E3A5F;
        position: relative;
        overflow: hidden;
        min-height: 94px;
        margin-bottom: 2px;
    }
    .dash-kpi.amber { border-top: 3px solid #BA7517; }
    .dash-kpi.lime  { border-top: 3px solid #C8E04A; }
    .dash-kpi.blue  { border-top: 3px solid #1B5179; }
    .dash-kpi.green { border-top: 3px solid #5A9F37; }
    .dash-kpi .dk-lbl {
        font-size: 10px;
        font-family: 'Inter', sans-serif;
        color: #8BA9C0;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        margin-bottom: 5px;
    }
    .dash-kpi .dk-val {
        font-family: 'Inter', sans-serif;
        font-size: 21px;
        font-weight: 600;
        color: #FFFFFF;
        font-feature-settings: "tnum";
        line-height: 1.15;
    }
    .dash-kpi .dk-delta {
        font-family: 'Inter', sans-serif;
        font-size: 11px;
        margin-top: 5px;
        color: #8BA9C0;
        min-height: 16px;
    }
    .dash-kpi .dk-spark {
        position: absolute;
        right: 10px;
        bottom: 8px;
        opacity: 0.7;
    }
    </style>
    """, unsafe_allow_html=True)

    _regs_dash = st.session_state.get("registros_acumulados", [])

    if not _regs_dash:
        st.markdown("""
        <div style="text-align:center; padding:80px 0; background:#0A2540; border-radius:14px;">
            <div style="font-size:52px; margin-bottom:18px;">📈</div>
            <div style="font-family:'Poppins',sans-serif; font-size:17px; font-weight:600; color:#FFFFFF; margin-bottom:10px;">
                Nenhuma fatura acumulada ainda
            </div>
            <div style="font-size:13px; color:#8BA9C0;">
                Processe faturas na aba <strong style="color:#C8E04A;">Processar faturas</strong> para ver o dashboard.
            </div>
        </div>
        """, unsafe_allow_html=True)
    else:
        _df = pd.DataFrame(_regs_dash).copy()
        _df["_data"]        = _df["ref_mes_ano"].apply(_parse_ref_dash)
        _df["total_fatura"] = pd.to_numeric(_df["total_fatura"],  errors="coerce").fillna(0)
        _df["__dif_total__"]= pd.to_numeric(_df["__dif_total__"], errors="coerce").fillna(0)
        _df["consumo_kwh"]  = pd.to_numeric(_df["consumo_kwh"],   errors="coerce").fillna(0)
        _df = _df.sort_values("_data")
        _df["_label"] = _df["_data"].dt.strftime("%m/%Y")

        # ── Helpers visuais ───────────────────────────────────────────────────
        def _spark(values, color="#C8E04A", w=88, h=28):
            """Gera sparkline SVG inline a partir de lista de valores."""
            vals = [float(v) for v in values if v is not None]
            if len(vals) < 2:
                return ""
            mn, mx = min(vals), max(vals)
            rng = mx - mn if mx != mn else 1
            pts = []
            n = len(vals)
            for i, v in enumerate(vals):
                x = i * w / (n - 1)
                y = h - 2 - (v - mn) / rng * (h - 4)
                pts.append(f"{x:.1f},{y:.1f}")
            poly = " ".join(pts)
            return (
                f'<svg width="{w}" height="{h}" viewBox="0 0 {w} {h}" '
                f'xmlns="http://www.w3.org/2000/svg">'
                f'<polyline points="{poly}" fill="none" stroke="{color}" '
                f'stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/></svg>'
            )

        def _delta_html(vals, invert=False):
            """Seta + % vs período anterior."""
            if len(vals) < 2 or vals[-2] == 0:
                return ""
            pct = (vals[-1] - vals[-2]) / abs(vals[-2]) * 100
            clr = ("#E05A5A" if pct > 0 else "#5A9F37") if invert else ("#5A9F37" if pct > 0 else "#E05A5A")
            arrow = "▲" if pct > 0 else "▼"
            return f'<span style="color:{clr};">{arrow} {abs(pct):.1f}% vs anterior</span>'

        # ── Constantes de tema dark ───────────────────────────────────────────
        _DBG = "#0D1B2A"    # paper bg
        _DCG = "#112233"    # plot bg
        _GRD = "#1E3A5F"    # grid
        _TXT = "#FFFFFF"
        _LBL = "#8BA9C0"
        _FF  = "Inter, sans-serif"
        _FFT = "Poppins, sans-serif"

        def _dl(title="", height=290, margin=None, showlegend=True, **kw):
            """Layout base dark para Plotly."""
            m = margin or dict(t=52, r=18, l=58, b=38)
            d = dict(
                paper_bgcolor=_DBG, plot_bgcolor=_DCG,
                font=dict(family=_FF, color=_TXT, size=12),
                title_text=title,
                title_font=dict(family=_FFT, color=_TXT, size=13),
                height=height, margin=m, showlegend=showlegend,
                legend=dict(font=dict(color=_LBL, size=11),
                            bgcolor="rgba(0,0,0,0)",
                            orientation="h", y=1.12, x=0),
                xaxis=dict(gridcolor=_GRD, linecolor=_GRD,
                           tickfont=dict(color=_LBL, size=11),
                           showgrid=True, zeroline=False),
                yaxis=dict(gridcolor=_GRD, linecolor=_GRD,
                           tickfont=dict(color=_LBL, size=11),
                           showgrid=True, zeroline=False),
            )
            d.update(kw)
            return d

        # ── Agregações globais ────────────────────────────────────────────────
        _total_gasto = _df["total_fatura"].sum()
        _val_rec     = _df["__dif_total__"].sum()
        _n_ucs       = _df["conta_uc"].nunique()
        _n_fat       = len(_df)
        _n_prob      = _df["__triagem__"].isin(["INVESTIGAR", "DIVERGENCIA"]).sum()
        _pct_ok      = ((_n_fat - _n_prob) / _n_fat * 100) if _n_fat else 0

        _order = (
            _df.dropna(subset=["_data"])
            .drop_duplicates("_label")
            .sort_values("_data")["_label"]
        )
        _mensal_s = (
            _df.dropna(subset=["_data"])
            .groupby("_label", sort=False)["total_fatura"].sum()
            .reindex(_order.values).fillna(0)
        )
        _mensal_rec_s = (
            _df.dropna(subset=["_data"])
            .groupby("_label", sort=False)["__dif_total__"].sum()
            .reindex(_order.values).fillna(0)
        )
        _mensal_prob_s = (
            _df[_df["__triagem__"].isin(["INVESTIGAR", "DIVERGENCIA"])]
            .dropna(subset=["_data"])
            .groupby("_label", sort=False).size()
            .reindex(_order.values).fillna(0)
        )
        _vmen  = list(_mensal_s.values)
        _vrec  = list(_mensal_rec_s.values)
        _vprob = list(_mensal_prob_s.values)

        # ── VISÃO GERAL ───────────────────────────────────────────────────────
        st.markdown('<div class="dash-section-title">Visão Geral</div>', unsafe_allow_html=True)

        _g1, _g2, _g3, _g4 = st.columns(4)
        with _g1:
            st.markdown(f"""
            <div class="dash-kpi amber">
              <div class="dk-lbl">Total Gasto no Período</div>
              <div class="dk-val">{_brl_d(_total_gasto)}</div>
              <div class="dk-delta">{_delta_html(_vmen, invert=True)}</div>
              <div class="dk-spark">{_spark(_vmen, "#BA7517")}</div>
            </div>""", unsafe_allow_html=True)
        with _g2:
            st.markdown(f"""
            <div class="dash-kpi lime">
              <div class="dk-lbl">Valor Passível de Recuperação</div>
              <div class="dk-val">{_brl_d(_val_rec)}</div>
              <div class="dk-delta">{_delta_html(_vrec)}</div>
              <div class="dk-spark">{_spark(_vrec, "#C8E04A")}</div>
            </div>""", unsafe_allow_html=True)
        with _g3:
            st.markdown(f"""
            <div class="dash-kpi blue">
              <div class="dk-lbl">Unidades Consumidoras</div>
              <div class="dk-val">{_n_ucs}</div>
              <div class="dk-delta"><span style="color:#8BA9C0;">{_n_fat} faturas processadas</span></div>
            </div>""", unsafe_allow_html=True)
        with _g4:
            _ok_clr = "#5A9F37" if _n_prob == 0 else "#BA7517"
            _prob_class = "green" if _n_prob == 0 else "amber"
            _prob_delta = (_delta_html(_vprob, invert=True)
                          or f'<span style="color:#5A9F37;">{_pct_ok:.0f}% em conformidade</span>')
            st.markdown(f"""
            <div class="dash-kpi {_prob_class}">
              <div class="dk-lbl">Faturas com Ocorrência</div>
              <div class="dk-val" style="color:{_ok_clr};">{_n_prob}<span style="font-size:13px;color:#8BA9C0;"> / {_n_fat}</span></div>
              <div class="dk-delta">{_prob_delta}</div>
              <div class="dk-spark">{_spark(_vprob, "#E05A5A")}</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

        if _PLOTLY_DASH:
            # Linha 1: gasto mensal (área) + donut triagem
            _ga, _gb = st.columns([3, 1])

            with _ga:
                _mensal_df = _mensal_s.reset_index()
                _mensal_df.columns = ["_label", "total_fatura"]
                _fig_men = _go.Figure(_go.Scatter(
                    x=_mensal_df["_label"], y=_mensal_df["total_fatura"],
                    mode="lines+markers",
                    line=dict(color="#C8E04A", width=2.5),
                    marker=dict(size=7, color="#C8E04A",
                                line=dict(color="#0A2540", width=2)),
                    fill="tozeroy",
                    fillcolor="rgba(200,224,74,0.10)",
                    hovertemplate="%{x}<br><b>R$ %{y:,.2f}</b><extra></extra>",
                ))
                _fig_men.update_layout(
                    **_dl("Gasto Total Mensal — todas as UCs",
                          height=300, showlegend=False),
                    yaxis_title="R$",
                )
                st.plotly_chart(_fig_men, use_container_width=True)

            with _gb:
                _t_cnt = _df["__triagem__"].value_counts()
                _t_clr = {"OK": "#5A9F37", "INVESTIGAR": "#1B5179",
                          "DIVERGENCIA": "#BA7517"}
                _fig_do = _go.Figure(_go.Pie(
                    labels=list(_t_cnt.index), values=list(_t_cnt.values),
                    hole=0.60,
                    marker=dict(
                        colors=[_t_clr.get(l, "#9AA8B7") for l in _t_cnt.index],
                        line=dict(color=_DBG, width=3),
                    ),
                    textinfo="percent",
                    textfont=dict(color=_TXT, size=11),
                    hovertemplate="%{label}: %{value} faturas<extra></extra>",
                ))
                _fig_do.update_layout(
                    paper_bgcolor=_DBG,
                    title_text="Triagem",
                    title_font=dict(family=_FFT, color=_TXT, size=13),
                    font=dict(family=_FF, color=_TXT),
                    height=300,
                    margin=dict(t=52, r=8, l=8, b=40),
                    showlegend=True,
                    legend=dict(font=dict(color=_LBL, size=10),
                                bgcolor="rgba(0,0,0,0)",
                                orientation="v", x=0.02, y=0.5),
                    annotations=[dict(
                        text=f"<b>{_pct_ok:.0f}%</b><br>ok",
                        x=0.5, y=0.5,
                        font=dict(size=16, color=_TXT, family=_FFT),
                        showarrow=False,
                    )],
                )
                st.plotly_chart(_fig_do, use_container_width=True)

            # Linha 2: ranking por UC
            _uc_rank = (
                _df.groupby("conta_uc")
                .agg(total=("total_fatura", "sum"),
                     recuperavel=("__dif_total__", "sum"))
                .reset_index()
                .sort_values("total", ascending=True)
            )
            _fig_uc = _go.Figure()
            _fig_uc.add_trace(_go.Bar(
                y=_uc_rank["conta_uc"], x=_uc_rank["total"],
                orientation="h", name="Total gasto",
                marker_color="#1B5179",
                hovertemplate="UC %{y}<br>Total: R$ %{x:,.2f}<extra></extra>",
            ))
            _fig_uc.add_trace(_go.Bar(
                y=_uc_rank["conta_uc"], x=_uc_rank["recuperavel"],
                orientation="h", name="Recuperável",
                marker_color="#BA7517",
                hovertemplate="UC %{y}<br>Recuperável: R$ %{x:,.2f}<extra></extra>",
            ))
            _fig_uc.update_layout(
                **_dl("Gasto Total e Valor Recuperável por UC",
                      height=max(220, len(_uc_rank) * 52 + 90),
                      margin=dict(t=52, r=18, l=95, b=38),
                      barmode="overlay"),
                xaxis_title="R$",
            )
            st.plotly_chart(_fig_uc, use_container_width=True)

        # ── POR UC ────────────────────────────────────────────────────────────
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
        st.markdown('<div class="dash-section-title">Por Unidade Consumidora</div>',
                    unsafe_allow_html=True)

        _ucs_d    = sorted(_df["conta_uc"].dropna().astype(str).unique())
        _uc_sel_d = st.selectbox("Selecionar UC", _ucs_d, key="dash_uc")
        _dfu      = _df[_df["conta_uc"].astype(str) == _uc_sel_d].copy()

        _uc_tot   = _dfu["total_fatura"].sum()
        _uc_rec   = _dfu["__dif_total__"].sum()
        _uc_kwh   = _dfu["consumo_kwh"].mean()
        _uc_np    = _dfu["__triagem__"].isin(["INVESTIGAR", "DIVERGENCIA"]).sum()
        _uc_nf    = len(_dfu)
        _uc_pctok = ((_uc_nf - _uc_np) / _uc_nf * 100) if _uc_nf else 0

        _dfu_s    = _dfu.sort_values("_data")
        _vuc_men  = list(_dfu_s["total_fatura"].values)
        _vuc_rec  = list(_dfu_s["__dif_total__"].values)
        _vuc_kwh  = list(_dfu_s["consumo_kwh"].values)
        _vuc_prob = list(
            _dfu_s["__triagem__"].isin(["INVESTIGAR", "DIVERGENCIA"])
            .astype(int).values
        )

        _u1, _u2, _u3, _u4 = st.columns(4)
        with _u1:
            st.markdown(f"""
            <div class="dash-kpi amber">
              <div class="dk-lbl">Total Gasto</div>
              <div class="dk-val">{_brl_d(_uc_tot)}</div>
              <div class="dk-delta">{_delta_html(_vuc_men, invert=True)}</div>
              <div class="dk-spark">{_spark(_vuc_men, "#BA7517")}</div>
            </div>""", unsafe_allow_html=True)
        with _u2:
            st.markdown(f"""
            <div class="dash-kpi lime">
              <div class="dk-lbl">Valor Recuperável</div>
              <div class="dk-val">{_brl_d(_uc_rec)}</div>
              <div class="dk-delta">{_delta_html(_vuc_rec)}</div>
              <div class="dk-spark">{_spark(_vuc_rec, "#C8E04A")}</div>
            </div>""", unsafe_allow_html=True)
        with _u3:
            st.markdown(f"""
            <div class="dash-kpi green">
              <div class="dk-lbl">Consumo Médio Mensal</div>
              <div class="dk-val">{_uc_kwh:,.0f}<span style="font-size:13px;color:#8BA9C0;"> kWh</span></div>
              <div class="dk-delta">{_delta_html(_vuc_kwh, invert=True)}</div>
              <div class="dk-spark">{_spark(_vuc_kwh, "#5A9F37")}</div>
            </div>""", unsafe_allow_html=True)
        with _u4:
            _uc_ok_clr   = "#5A9F37" if _uc_np == 0 else "#BA7517"
            _uc_prob_cls = "green" if _uc_np == 0 else "amber"
            _uc_prob_dlt = (_delta_html(_vuc_prob, invert=True)
                           or f'<span style="color:#5A9F37;">{_uc_pctok:.0f}% em conformidade</span>')
            st.markdown(f"""
            <div class="dash-kpi {_uc_prob_cls}">
              <div class="dk-lbl">Faturas com Ocorrência</div>
              <div class="dk-val" style="color:{_uc_ok_clr};">{_uc_np}<span style="font-size:13px;color:#8BA9C0;"> / {_uc_nf}</span></div>
              <div class="dk-delta">{_uc_prob_dlt}</div>
              <div class="dk-spark">{_spark(_vuc_prob, "#E05A5A")}</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

        if _PLOTLY_DASH:
            _dem_med_s  = _col(_dfu_s, "demanda_medida_kw")
            _dem_cont_s = _col(_dfu_s, "demanda_contratada_kw")
            _has_demand = _dem_med_s.sum() > 0 and _dem_cont_s.sum() > 0

            # Gasto mensal colorido por triagem
            _tri_clrs = [
                "#BA7517" if t in ("INVESTIGAR", "DIVERGENCIA") else "#1B5179"
                for t in _dfu_s["__triagem__"]
            ]
            _fig_fat = _go.Figure(_go.Bar(
                x=_dfu_s["_label"], y=_dfu_s["total_fatura"],
                marker_color=_tri_clrs,
                hovertemplate="%{x}<br><b>R$ %{y:,.2f}</b><extra></extra>",
            ))
            _fig_fat.update_layout(
                **_dl(f"Evolução do Gasto Mensal — UC {_uc_sel_d}",
                      height=270, showlegend=False,
                      margin=dict(t=52, r=18, l=58, b=38)),
                yaxis_title="R$",
            )

            if _has_demand:
                # Gauge de utilização de demanda + gasto side-by-side
                _dem_pico  = _dem_med_s.max()
                _dem_ctrt  = _dem_cont_s.iloc[-1]
                _dem_pct   = (_dem_pico / _dem_ctrt * 100) if _dem_ctrt else 0
                _dem_pct_c = min(_dem_pct, 150)
                _g_clr = ("#5A9F37" if _dem_pct <= 100
                          else "#BA7517" if _dem_pct <= 110
                          else "#E05A5A")

                _col_ga, _col_gb = st.columns([1, 2])
                with _col_ga:
                    _fig_gauge = _go.Figure(_go.Indicator(
                        mode="gauge+number+delta",
                        value=_dem_pct_c,
                        number=dict(suffix="%",
                                    font=dict(size=26, color=_TXT, family=_FF)),
                        delta=dict(reference=100,
                                   increasing=dict(color="#E05A5A"),
                                   decreasing=dict(color="#5A9F37"),
                                   font=dict(size=12)),
                        gauge=dict(
                            axis=dict(range=[0, 150],
                                      tickcolor=_LBL,
                                      tickfont=dict(color=_LBL, size=10),
                                      nticks=6),
                            bar=dict(color=_g_clr, thickness=0.22),
                            bgcolor=_DCG, borderwidth=0,
                            steps=[
                                dict(range=[0, 95],   color="#0F2A10"),
                                dict(range=[95, 110],  color="#2A1F00"),
                                dict(range=[110, 150], color="#2A0A0A"),
                            ],
                            threshold=dict(
                                line=dict(color="#FFFFFF", width=2),
                                thickness=0.8, value=100,
                            ),
                        ),
                        title=dict(
                            text=(f"Utilização de Demanda<br>"
                                  f"<span style='font-size:10px;color:{_LBL};'>"
                                  f"Pico {_dem_pico:.1f} kW · "
                                  f"Contrat. {_dem_ctrt:.0f} kW</span>"),
                            font=dict(size=12, color=_TXT, family=_FFT),
                        ),
                    ))
                    _fig_gauge.update_layout(
                        paper_bgcolor=_DBG, height=270,
                        margin=dict(t=62, r=18, l=18, b=18),
                        font=dict(color=_TXT),
                    )
                    st.plotly_chart(_fig_gauge, use_container_width=True)

                with _col_gb:
                    st.plotly_chart(_fig_fat, use_container_width=True)
            else:
                st.plotly_chart(_fig_fat, use_container_width=True)

            _col_e, _col_f = st.columns(2)

            with _col_e:
                _kp  = _col(_dfu_s, "consumo_ponta_kwh")
                _kfp = _col(_dfu_s, "consumo_fp_kwh")
                if _kp.sum() + _kfp.sum() > 0:
                    _fig_kwh = _go.Figure()
                    _fig_kwh.add_trace(_go.Bar(
                        x=_dfu_s["_label"], y=_kp, name="Ponta",
                        marker_color="#BA7517",
                        hovertemplate="%{x}<br>Ponta: %{y:,.0f} kWh<extra></extra>",
                    ))
                    _fig_kwh.add_trace(_go.Bar(
                        x=_dfu_s["_label"], y=_kfp, name="Fora Ponta",
                        marker_color="#1B5179",
                        hovertemplate="%{x}<br>F.Ponta: %{y:,.0f} kWh<extra></extra>",
                    ))
                    _fig_kwh.update_layout(
                        **_dl("Consumo kWh — Ponta vs Fora Ponta",
                              height=260, barmode="stack",
                              margin=dict(t=52, r=14, l=54, b=38)),
                        yaxis_title="kWh",
                    )
                else:
                    _fig_kwh = _go.Figure(_go.Bar(
                        x=_dfu_s["_label"], y=_dfu_s["consumo_kwh"],
                        marker_color="#1B5179",
                        hovertemplate="%{x}<br>%{y:,.0f} kWh<extra></extra>",
                    ))
                    _fig_kwh.update_layout(
                        **_dl("Consumo kWh", height=260, showlegend=False,
                              margin=dict(t=52, r=14, l=54, b=38)),
                        yaxis_title="kWh",
                    )
                st.plotly_chart(_fig_kwh, use_container_width=True)

            with _col_f:
                _dem_med  = _col(_dfu_s, "demanda_medida_kw")
                _dem_cont = _col(_dfu_s, "demanda_contratada_kw")
                if _dem_med.sum() > 0:
                    _fig_dem = _go.Figure()
                    _fig_dem.add_trace(_go.Bar(
                        x=_dfu_s["_label"], y=_dem_med,
                        name="Medida", marker_color="#1B5179",
                        hovertemplate="%{x}<br>Medida: %{y:,.1f} kW<extra></extra>",
                    ))
                    if _dem_cont.sum() > 0:
                        _fig_dem.add_trace(_go.Scatter(
                            x=_dfu_s["_label"], y=_dem_cont,
                            mode="lines+markers", name="Contratada",
                            line=dict(color="#C8E04A", width=2, dash="dash"),
                            marker=dict(size=5, color="#C8E04A"),
                            hovertemplate="%{x}<br>Contratada: %{y:,.1f} kW<extra></extra>",
                        ))
                    _fig_dem.update_layout(
                        **_dl("Demanda kW — Medida vs Contratada",
                              height=260, margin=dict(t=52, r=14, l=54, b=38)),
                        yaxis_title="kW",
                    )
                    st.plotly_chart(_fig_dem, use_container_width=True)
                else:
                    _icms = _col(_dfu_s, "icms_valor")
                    _pis  = _col(_dfu_s, "pis_valor")
                    _cof  = _col(_dfu_s, "cofins_valor")
                    _fig_trib = _go.Figure()
                    _fig_trib.add_trace(_go.Bar(
                        x=_dfu_s["_label"], y=_icms, name="ICMS",
                        marker_color="#0A2540"))
                    _fig_trib.add_trace(_go.Bar(
                        x=_dfu_s["_label"], y=_pis, name="PIS",
                        marker_color="#1B5179"))
                    _fig_trib.add_trace(_go.Bar(
                        x=_dfu_s["_label"], y=_cof, name="COFINS",
                        marker_color="#5A9F37"))
                    _fig_trib.update_layout(
                        **_dl("Tributos por Mês", height=260, barmode="stack",
                              margin=dict(t=52, r=14, l=54, b=38)),
                        yaxis_title="R$",
                    )
                    st.plotly_chart(_fig_trib, use_container_width=True)

            # Custos acessórios (reativo + USDG ultrap)
            _rea_p  = _col(_dfu_s, "valor_reativo_exc_ponta")
            _rea_fp = _col(_dfu_s, "valor_reativo_exc_fp")
            _usdg_u = _col(_dfu_s, "valor_usdg_ultrap")
            _total_aces = _rea_p.sum() + _rea_fp.sum() + _usdg_u.sum()
            if _total_aces > 0:
                _fig_aces = _go.Figure()
                if (_rea_p + _rea_fp).sum() > 0:
                    _fig_aces.add_trace(_go.Bar(
                        x=_dfu_s["_label"], y=(_rea_p + _rea_fp).values,
                        name="Reativo Exc.", marker_color="#BA7517",
                        hovertemplate="%{x}<br>Reativo: R$ %{y:,.2f}<extra></extra>",
                    ))
                if _usdg_u.sum() > 0:
                    _fig_aces.add_trace(_go.Bar(
                        x=_dfu_s["_label"], y=_usdg_u.values,
                        name="USDG Ultrap.", marker_color="#0A2540",
                        hovertemplate="%{x}<br>USDG Ultrap: R$ %{y:,.2f}<extra></extra>",
                    ))
                _fig_aces.update_layout(
                    **_dl("Custos Acessórios — Reativo Excedente + USDG Ultrapassagem",
                          height=260, barmode="stack",
                          margin=dict(t=52, r=18, l=64, b=38)),
                    yaxis_title="R$",
                )
                st.plotly_chart(_fig_aces, use_container_width=True)


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

        if _RELATORIO_PDF_ERRO:
            st.error(f"⚠️ Módulo PDF falhou ao carregar: `{_RELATORIO_PDF_ERRO}`")

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

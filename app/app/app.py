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

# Neoenergia PE — importação condicional
try:
    from neoenergia_pe.extractor import parse_fatura as neo_parse
    from neoenergia_pe.audit import auditar as neo_auditar
    NEO_DISPONIVEL = True
except ImportError:
    NEO_DISPONIVEL = False

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
        "cosip":          None,
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
</style>
""", unsafe_allow_html=True)


# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="me-header">
    <h1>⚡ Minuto Energia — Auditor de Faturas</h1>
    <p>Auditoria automática conforme REH ANEEL e Lei 14.300/2022</p>
</div>
""", unsafe_allow_html=True)


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
    st.markdown("### Configuração do Lote")
    st.markdown("---")

    opcoes_dist = ["CPFL Piratininga", "Neoenergia PE"]
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
tab1, tab2, tab3 = st.tabs(["📤 Processar faturas", "📊 Resultados", "ℹ️ Sobre"])


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
                    else:
                        res = processar_fatura(tmp_pdf, config_default, run_dir)
                    resultados.append(res)
                except Exception as e:
                    erros.append({"file": file.name, "error": str(e)})
                progress.progress((i + 1) / len(uploaded))

            status.markdown("📊 Consolidando resultados...")
            master_path = run_dir / f"MASTER_Auditoria_{run_id}.xlsx"
            registros_norm = [r["__registro__"] for r in resultados]
            gerar_excel_mestre(registros_norm, master_path)

            status.markdown(f"✅ **Auditoria conferida!** {len(resultados)} fatura(s) processada(s).")

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
                "Ref Mês/Ano": reg.get("ref_mes_ano", ""),
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


# ── Tab 3: Sobre ──────────────────────────────────────────────────────────────
with tab3:
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
- Fio B — Lei 14.300/2022 (pré-MM

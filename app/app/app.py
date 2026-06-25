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
from master_excel import gerar_master

# Neoenergia PE — importação condicional
try:
    from neoenergia_pe.extractor import parse_fatura as neo_parse
    from neoenergia_pe.audit import auditar as neo_auditar
    NEO_DISPONIVEL = True
except ImportError:
    NEO_DISPONIVEL = False

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

    # Converte para formato compatível com a UI
    status_map = {"OK": "OK", "INVESTIGAR": "INVESTIGAR", "DIVERGENCIA": "ATENÇÃO"}
    status = status_map.get(triagem, "OK")

    if motivos:
        alertas = [{"cat": "Auditoria", "descricao": m, "status": status} for m in motivos]
    else:
        alertas = [{"cat": "Auditoria", "descricao": "Fatura conferida", "status": "OK"}]

    dados = {
        "cliente_nome":    r.get("conta_contrato", ""),
        "uc":              r.get("conta_contrato", ""),
        "subgrupo":        "B3",
        "leitura_anterior": None,
        "leitura_atual":   None,
        "dias_ciclo":      r.get("nr_dias"),
        "total_fatura":    r.get("total_fatura"),
        "total_a_pagar":   r.get("total_fatura"),
        "mes_ref":         r.get("ref_mes_ano", ""),
        "nota_fiscal":     r.get("nr_nota_fiscal", ""),
    }

    return {
        "pdf_filename": pdf_path.name,
        "dados": dados,
        "config": {},
        "audit_result": {"alertas": alertas, "auditado": {"reh_aplicada": ""}},
        "excel_individual_path": "",
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

    return {
        "pdf_filename": pdf_path.name,
        "dados": dados,
        "config": config,
        "audit_result": audit_result,
        "excel_individual_path": str(out_xlsx),
    }


with st.sidebar:
    st.markdown("### Configuração do Lote")
    st.markdown("---")

    opcoes_dist = ["CPFL Piratininga", "Neoenergia PE"]
    distribuidora = st.selectbox("Distribuidora", opcoes_dist)

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

    uploaded = st.file_uploader(
        "Arraste os PDFs aqui ou clique para selecionar",
        type=["pdf"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    if uploaded:
        st.markdown(f"**{len(uploaded)} fatura(s) selecionada(s)**")
        cols = st.columns(3)
        for i, u in enumerate(uploaded):
            cols[i % 3].markdown(f"📄 {u.name} · {u.size // 1024} KB")

        st.markdown("")
        if st.button(f"Processar {len(uploaded)} fatura(s)", type="primary"):
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
            gerar_master(resultados, master_path)

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

        total_ok   = sum(summary_alertas(r["audit_result"]["alertas"])["OK"] for r in resultados)
        total_aten = sum(summary_alertas(r["audit_result"]["alertas"])["ATENÇÃO"] for r in resultados)
        total_inv  = sum(summary_alertas(r["audit_result"]["alertas"])["INVESTIGAR"] for r in resultados)
        total_valor = sum(r["dados"].get("total_fatura") or 0 for r in resultados)
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
            <div class="lbl">Para investigação</div>
        </div>""", unsafe_allow_html=True)
        c4.markdown(f"""
        <div class="me-metric div">
            <div class="val">{total_aten}</div>
            <div class="lbl">Atenção necessária</div>
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
            d = r["dados"]
            counts = summary_alertas(r["audit_result"]["alertas"])
            n_inv = counts["INVESTIGAR"]
            n_at  = counts["ATENÇÃO"]
            if n_inv > 0:
                status_html = f'<span class="badge badge-inv">INVESTIGAR ({n_inv})</span>'
            elif n_at > 0:
                status_html = f'<span class="badge badge-div">ATENÇÃO ({n_at})</span>'
            else:
                status_html = '<span class="badge badge-ok">CONFERIDO</span>'
            rows.append({
                "Arquivo": r["pdf_filename"],
                "Cliente": (d.get("cliente_nome", "") or "")[:30],
                "UC": d.get("uc", ""),
                "Período": f"{d.get('mes_ref', '')}",
                "Dias": d.get("dias_ciclo", ""),
                "Total (R$)": d.get("total_fatura"),
                "Conferidos": counts["OK"],
                "Investigar": counts["INVESTIGAR"],
                "Atenção": counts["ATENÇÃO"],
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
            d = r["dados"]
            counts = summary_alertas(r["audit_result"]["alertas"])
            label = f"{r['pdf_filename']} · {(d.get('cliente_nome','') or '')[:30]} · UC {d.get('uc', '')}"
            badge = ""
            if counts["ATENÇÃO"] > 0:
                badge += f" ⚠️ {counts['ATENÇÃO']}"
            if counts["INVESTIGAR"] > 0:
                badge += f" 🔍 {counts['INVESTIGAR']}"
            with st.expander(label + badge):
                for a in r["audit_result"]["alertas"]:
                    icon = {"OK": "✅", "ATENÇÃO": "⚠️", "INVESTIGAR": "🔍"}.get(a["status"], "❓")
                    st.markdown(f"{icon} **[{a['cat']}]** {a['descricao']}")


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
- Fio B — Lei 14.300/2022 (pré-MMGD isento, pós-MMGD escalonado)
- Período de leitura dentro dos limites da REN 1.000/2021 (15–45 dias)
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
    st.caption("Minuto Energia · Gestão e Eficiência Energética · minutoenergia.com.br")

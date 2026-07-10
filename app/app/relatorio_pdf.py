# -*- coding: utf-8 -*-
"""relatorio_pdf.py — Minuto Energia"""
from __future__ import annotations
import io
from collections import defaultdict

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Spacer, Image, Table, TableStyle,
    Paragraph, HRFlowable, PageBreak, Flowable,
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor

NAVY    = HexColor("#0A2540")
BLUE    = HexColor("#1B5179")
GREEN   = HexColor("#5A9F37")
LIME    = HexColor("#C8E04A")
AMBER   = HexColor("#BA7517")
OFFWHITE= HexColor("#FAFBF8")
GRAY    = HexColor("#5A6B7C")
LGRAY   = HexColor("#E5EBE0")
WHITE   = colors.white

C_NAVY  = "#0A2540"; C_BLUE  = "#1B5179"; C_GREEN = "#5A9F37"
C_LIME  = "#C8E04A"; C_AMBER = "#BA7517"; C_GRAY  = "#5A6B7C"

PW, PH = landscape(A4)
MARGIN  = 15 * mm
CW      = PW - 2 * MARGIN

W_ALL = 250
H_ALL = 175

_MESES = {
    "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4,
    "MAI": 5, "JUN": 6, "JUL": 7, "AGO": 8,
    "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12,
}

def _brl(v):
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def _parse_ref(ref):
    try:
        partes = ref.upper().split("/")
        return (int(partes[1]), _MESES.get(partes[0], 0))
    except Exception:
        return (9999, 0)

def _periodo_anos(refs):
    if not refs:
        return "—"
    anos = sorted({_parse_ref(r)[0] for r in refs if r})
    if len(anos) == 1:
        return str(anos[0])
    return f"{anos[0]}–{anos[-1]}"

def _periodo_geral(refs):
    return _periodo_anos(refs)

def _fig_to_img(fig, dpi=200):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight",
                facecolor=fig.get_facecolor())
    buf.seek(0)
    plt.close(fig)
    return buf

FSIZE = (5.8, 4.0)

def _chart_donut(tot_ok, tot_inv):
    fig, ax = plt.subplots(figsize=FSIZE, facecolor="white")
    vals   = [tot_ok, tot_inv]
    clrs   = [C_GREEN, C_BLUE]
    labels = [f"OK ({tot_ok})", f"DIVERGENTE ({tot_inv})"]
    ax.pie(vals, colors=clrs, startangle=90,
           wedgeprops=dict(width=0.52, edgecolor="white", linewidth=2))
    ax.text(0, 0, f"{sum(vals)}\nFaturas", ha="center", va="center",
            fontsize=14, fontweight="bold", color=C_NAVY)
    patches = [mpatches.Patch(color=c, label=l) for c, l in zip(clrs, labels)]
    ax.legend(handles=patches, loc="center right", bbox_to_anchor=(1.45, 0.5),
              fontsize=8, frameon=False)
    ax.set_title("Distribuição das Faturas", fontsize=11, fontweight="bold",
                 color=C_NAVY, pad=8)
    fig.tight_layout()
    return _fig_to_img(fig)

def _chart_barras(rs_ok, rs_inv):
    fig, ax = plt.subplots(figsize=FSIZE, facecolor="white")
    cats = ["OK", "DIVERGENTE"]
    vals = [rs_ok, rs_inv]
    clrs = [C_GREEN, C_BLUE]
    x = np.arange(len(cats))
    bars = ax.bar(x, vals, color=clrs, width=0.45, edgecolor="white", linewidth=1.5)
    for bar, v in zip(bars, vals):
        label = f"R$ {v:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max(vals) * 0.02,
                label, ha="center", va="bottom", fontsize=8.5, color=C_NAVY, fontweight="bold")
    ax.set_xticks(x)
    ax.set_xticklabels(cats, fontsize=9)
    ax.set_ylabel("R$", fontsize=9, color=C_GRAY)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"R$ {v/1000:.0f}k"))
    ax.spines[["top", "right"]].set_visible(False)
    ax.spines[["left", "bottom"]].set_color("#E5EBE0")
    ax.set_title("Valor Total por Triagem", fontsize=11, fontweight="bold", color=C_NAVY, pad=8)
    ax.set_ylim(0, max(vals) * 1.28 if max(vals) > 0 else 1)
    fig.tight_layout()
    return _fig_to_img(fig)

def _chart_ucbar(ucs_data):
    fig, ax = plt.subplots(figsize=FSIZE, facecolor="white")
    labels = [f"...{u[0][-4:]}" if len(u[0]) > 4 else u[0] for u in ucs_data]
    vals   = [u[1] for u in ucs_data]
    y = np.arange(len(labels))
    ax.barh(y, vals, color=C_BLUE, height=0.52, label="DIVERGENTE")
    ax.set_yticks(y)
    ax.set_yticklabels(labels, fontsize=8.5)
    ax.set_xlabel("Nº de Faturas", fontsize=9)
    ax.spines[["top", "right"]].set_visible(False)
    ax.spines[["left", "bottom"]].set_color("#E5EBE0")
    ax.legend(fontsize=8.5, frameon=False, loc="lower right")
    ax.set_title("Faturas com Ocorrências por UC", fontsize=11, fontweight="bold",
                 color=C_NAVY, pad=8)
    fig.tight_layout()
    return _fig_to_img(fig)

def _draw_capa(c, cliente_nome, dist, periodo, n_ucs, data_rel, parceiro_nome=""):
    c.setFillColor(NAVY)
    c.rect(0, 0, PW, PH, fill=1, stroke=0)
    c.setFillColor(LIME)
    c.rect(0, 0, 6 * mm, PH, fill=1, stroke=0)
    c.setFillColor(BLUE)
    c.rect(6 * mm, 0, PW - 6 * mm, 22 * mm, fill=1, stroke=0)
    if parceiro_nome:
        c.setFillColor(LIME)
        c.setFont("Helvetica-Bold", 12)
        c.drawString(22 * mm, PH - 28 * mm, parceiro_nome.upper()[:60])
        c.setFont("Helvetica-Bold", 12)
        c.drawString(22 * mm, PH - 40 * mm, "MINUTO ENERGIA")
        c.setFillColor(WHITE)
        c.setFont("Helvetica", 9.5)
        c.drawString(22 * mm, PH - 50 * mm, "Gestão e Eficiência Energética")
    else:
        c.setFillColor(LIME)
        c.setFont("Helvetica-Bold", 12)
        c.drawString(22 * mm, PH - 32 * mm, "MINUTO ENERGIA")
        c.setFillColor(WHITE)
        c.setFont("Helvetica", 9.5)
        c.drawString(22 * mm, PH - 42 * mm, "Gestão e Eficiência Energética")
    c.setStrokeColor(LIME)
    c.setLineWidth(1.5)
    c.line(22 * mm, PH - 58 * mm, 160 * mm, PH - 58 * mm)
    c.setFillColor(WHITE)
    c.setFont("Helvetica-Bold", 32)
    c.drawString(22 * mm, PH - 90 * mm, "AUDITORIA DE")
    c.drawString(22 * mm, PH - 124 * mm, "FATURAS DE ENERGIA")
    c.setFillColor(LIME)
    c.setFont("Helvetica-Bold", 13)
    c.drawString(22 * mm, PH - 148 * mm, cliente_nome[:70])
    c.setFillColor(WHITE)
    c.setFont("Helvetica", 10)
    c.drawString(22 * mm, PH - 163 * mm,
                 f"{dist}   |   Período: {periodo}   |   {n_ucs} Unidade(s) Consumidora(s)")
    c.setFillColor(LGRAY)
    c.setFont("Helvetica", 9)
    c.drawString(22 * mm, 9 * mm, data_rel)
    c.drawRightString(PW - 14 * mm, 9 * mm, "minutoenergia.com.br")

def _draw_contracapa(c, parceiro_nome=""):
    c.setFillColor(NAVY)
    c.rect(0, 0, PW, PH, fill=1, stroke=0)
    c.setFillColor(LIME)
    c.rect(0, 0, 6 * mm, PH, fill=1, stroke=0)
    c.setFillColor(BLUE)
    c.rect(6 * mm, 0, PW - 6 * mm, 22 * mm, fill=1, stroke=0)
    c.setFillColor(WHITE)
    c.setFont("Helvetica-Bold", 19)
    c.drawCentredString(PW / 2, PH / 2 + 12 * mm,
                        "Conte conosco na defesa dos seus interesses,")
    c.drawCentredString(PW / 2, PH / 2 - 4 * mm,
                        "sempre com transparência e efetividade!")
    c.setFillColor(LIME)
    c.setFont("Helvetica-Bold", 12)
    if parceiro_nome:
        c.drawCentredString(PW / 2, PH / 2 - 22 * mm, parceiro_nome.upper()[:60])
        c.drawCentredString(PW / 2, PH / 2 - 36 * mm, "MINUTO ENERGIA")
    else:
        c.drawCentredString(PW / 2, PH / 2 - 22 * mm, "MINUTO ENERGIA")
    c.setFillColor(LGRAY)
    c.setFont("Helvetica", 9)
    c.drawCentredString(PW / 2, 9 * mm, "minutoenergia.com.br")


class _TriggerContracapa(Flowable):
    """Marcador zero-size: sinaliza ao canvas que esta página é a contracapa."""
    width = 0
    height = 0
    def draw(self):
        self.canv._contracapa_this_page = True


class _RelatorioCanvas(canvas.Canvas):
    def __init__(self, filename, cliente_nome="", parceiro_nome="", **kwargs):
        super().__init__(filename, **kwargs)
        self._page_num             = 0
        self._cliente_nome         = cliente_nome
        self._parceiro_nome        = parceiro_nome
        self._contracapa_this_page = False

    def showPage(self):
        self._page_num += 1
        if self._page_num == 1:
            pass  # capa desenhada por _on_first_page
        elif self._contracapa_this_page:
            _draw_contracapa(self, self._parceiro_nome)
            self._contracapa_this_page = False
        else:
            self._draw_header_footer(self._page_num)
        super().showPage()

    def save(self):
        super().save()

    def _draw_header_footer(self, pn):
        self.setFillColor(NAVY)
        self.rect(0, PH - 18 * mm, PW, 18 * mm, fill=1, stroke=0)
        if self._parceiro_nome:
            self.setFillColor(LIME)
            self.setFont("Helvetica-Bold", 8)
            self.drawString(12 * mm, PH - 7 * mm, self._parceiro_nome.upper()[:50])
            self.setFont("Helvetica-Bold", 8)
            self.drawString(12 * mm, PH - 13.5 * mm, "MINUTO ENERGIA  |  Auditoria de Faturas")
        else:
            self.setFillColor(LIME)
            self.setFont("Helvetica-Bold", 9)
            self.drawString(12 * mm, PH - 11 * mm, "MINUTO ENERGIA  |  Auditoria de Faturas")
        self.setFillColor(WHITE)
        self.setFont("Helvetica", 8)
        self.drawRightString(PW - 12 * mm, PH - 11 * mm, self._cliente_nome[:60])
        self.setFillColor(LGRAY)
        self.rect(0, 0, PW, 8 * mm, fill=1, stroke=0)
        self.setFillColor(GRAY)
        self.setFont("Helvetica", 7.5)
        self.drawString(12 * mm, 3 * mm,
                        "Minuto Energia — Gestão e Eficiência Energética — minutoenergia.com.br")
        self.drawRightString(PW - 12 * mm, 3 * mm, f"Pág. {pn - 1}")


def _S(name, **kw):
    return ParagraphStyle(name, **kw)

_ST = {}

def _get_styles():
    if _ST:
        return _ST
    _ST["section"] = _S("sec", fontName="Helvetica-Bold", fontSize=16,
                         textColor=NAVY, spaceAfter=6, spaceBefore=4)
    _ST["body"]    = _S("body", fontName="Helvetica", fontSize=10,
                         textColor=GRAY, leading=14, spaceAfter=4)
    _ST["bold"]    = _S("bold", fontName="Helvetica-Bold", fontSize=10,
                         textColor=NAVY, leading=14)
    _ST["prazo_d"] = _S("pd",   fontName="Helvetica", fontSize=9,
                         textColor=GRAY, leading=13, spaceAfter=8)
    _ST["obs"]     = _S("obs",  fontName="Helvetica-Oblique", fontSize=8.5,
                         textColor=GRAY, leading=12)
    return _ST

PRAZOS = [
    ("Protocolo inicial",
     "Após assinatura do contrato, o protocolo de contestação é realizado em até 10 dias úteis."),
    ("Resposta da distribuidora",
     "Prazo inicial de 7 dias úteis para análise, prorrogável por até 30 dias."),
    ("Ouvidoria da distribuidora",
     "Contestação em até 7 dias úteis. Prazo de resposta: 15 dias, prorrogável por 60 dias."),
    ("Ouvidoria ANEEL",
     "Contestação em até 7 dias úteis. Prazo de resposta: 15 dias, prorrogável por 90 dias."),
    ("SMA — Superintendência ANEEL",
     "Contestação em até 10 dias úteis. A ANEEL não possui prazo legal de resposta nesta etapa."),
]
OBS_PRAZOS = (
    "* Para que toda a tratativa seja possível, é necessário que o cliente "
    "assine o contrato e a procuração com poderes específicos."
)

def gerar_relatorio_pdf(cliente_nome, registros, parceiro_nome=""):
    import datetime as _dt
    ST = _get_styles()

    ucs_dict = defaultdict(
        lambda: {"fat": 0, "ok": 0, "inv": 0, "div": 0, "total": 0.0, "refs": []}
    )
    for reg in registros:
        uc = reg.get("conta_uc") or "—"
        t  = reg.get("__triagem__", "OK")
        ucs_dict[uc]["fat"]   += 1
        ucs_dict[uc]["ok"]    += 1 if t == "OK" else 0
        ucs_dict[uc]["inv"]   += 1 if t == "INVESTIGAR" else 0
        ucs_dict[uc]["div"]   += 1 if t == "DIVERGENCIA" else 0
        ucs_dict[uc]["total"] += reg.get("total_fatura") or 0.0
        ref = reg.get("ref_mes_ano")
        if ref:
            ucs_dict[uc]["refs"].append(ref)

    UCS = [
        (uc, _periodo_anos(d["refs"]), d["fat"], d["ok"], d["inv"], d["div"], d["total"])
        for uc, d in sorted(ucs_dict.items())
    ]

    TOT_OK   = sum(d["ok"]  for d in ucs_dict.values())
    TOT_INV  = sum(d["inv"] for d in ucs_dict.values())
    TOT_DIV  = sum(d["div"] for d in ucs_dict.values())
    RS_OK    = sum(r.get("total_fatura") or 0 for r in registros if r.get("__triagem__") == "OK")
    RS_INV   = sum(r.get("total_fatura") or 0 for r in registros if r.get("__triagem__") == "INVESTIGAR")
    RS_DIV   = sum(r.get("total_fatura") or 0 for r in registros if r.get("__triagem__") == "DIVERGENCIA")
    RS_GERAL = sum(r.get("total_fatura") or 0 for r in registros)
    N_FAT    = len(registros)
    N_UCS    = len(ucs_dict)
    TOT_PROB = TOT_INV + TOT_DIV
    RS_PROB  = RS_INV + RS_DIV

    dists   = sorted({r.get("distribuidora") for r in registros if r.get("distribuidora")})
    DIST    = " / ".join(dists) if dists else "—"
    all_refs = [r.get("ref_mes_ano") for r in registros if r.get("ref_mes_ano")]
    PERIODO  = _periodo_geral(all_refs)
    DATA_REL = _dt.date.today().strftime("%B / %Y").capitalize()

    donut_buf  = _chart_donut(TOT_OK, TOT_PROB)
    barras_buf = _chart_barras(RS_OK, RS_PROB)
    ucbar_buf  = _chart_ucbar([(uc, d["inv"] + d["div"]) for uc, d in ucs_dict.items()])

    story = []
    story.append(PageBreak())   # p.1 = capa

    # P2: Dados da Análise
    story.append(Paragraph("Dados da Análise", ST["section"]))
    story.append(HRFlowable(width="100%", thickness=1.5, color=LIME, spaceAfter=10))

    metrics = [
        ("Distribuidora",        DIST,    C_BLUE),
        ("Período Analisado",    PERIODO, C_NAVY),
        ("Unidades Consumidoras",str(N_UCS), C_GREEN),
        ("Faturas Auditadas",    str(N_FAT), "#5A6B7C"),
    ]

    def _metric_card(label, value, color):
        return Table(
            [[Paragraph(f'<font color="{color}"><b>{value}</b></font>',
                        ParagraphStyle("mv", fontName="Helvetica-Bold", fontSize=22,
                                       textColor=HexColor(color), alignment=1))],
             [Paragraph(label, ParagraphStyle("ml", fontName="Helvetica", fontSize=9,
                                              textColor=GRAY, alignment=1))]],
            colWidths=[(CW - 24) / 4],
            style=TableStyle([
                ("BOX",           (0, 0), (-1, -1), 1,  LGRAY),
                ("TOPPADDING",    (0, 0), (-1, -1), 10),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
                ("BACKGROUND",    (0, 0), (-1, -1), WHITE),
            ]))

    story.append(Table(
        [[_metric_card(l, v, c) for l, v, c in metrics]],
        colWidths=[CW / 4] * 4,
        style=TableStyle([
            ("LEFTPADDING",  (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ])))
    story.append(Spacer(1, 8 * mm))

    pontos = [
        "Tarifas TUSD e TE (REH ANEEL vigente)",
        "Tributos: ICMS, PIS, COFINS — alíquotas e bases de cálculo",
        "Bandeira tarifária — proporcionalidade por período",
        "Leituras de consumo — consistência medidor vs faturado",
        "Período de leitura — limites regulatórios (15–45 dias, REN 1.000/2021)",
        "Cobranças retroativas (multas, juros, atualizações)",
        "Divergência entre total da fatura e total a pagar",
        "GD — compensação de energia injetada",
    ]
    story.append(Paragraph("Itens verificados na auditoria:", ST["bold"]))
    story.append(Spacer(1, 3 * mm))
    half = len(pontos) // 2
    col1 = [Paragraph(f"• {p}", ST["body"]) for p in pontos[:half]]
    col2 = [Paragraph(f"• {p}", ST["body"]) for p in pontos[half:]]
    mr = max(len(col1), len(col2))
    while len(col1) < mr:
        col1.append(Spacer(1, 1))
    while len(col2) < mr:
        col2.append(Spacer(1, 1))
    story.append(Table(
        [[col1[i], col2[i]] for i in range(mr)],
        colWidths=[CW / 2] * 2,
        style=TableStyle([
            ("TOPPADDING",    (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ])))
    story.append(PageBreak())

    # P3: Resultados Consolidados
    story.append(Paragraph("Resultados Consolidados", ST["section"]))
    story.append(HRFlowable(width="100%", thickness=1.5, color=LIME, spaceAfter=8))

    kpis = [
        (str(TOT_OK),    "faturas OK",          C_GREEN),
        (str(TOT_PROB),  "faturas Divergentes",  C_BLUE),
        (_brl(RS_PROB),  "em faturas com ocorrências", C_NAVY),
        (_brl(RS_GERAL), "valor total auditado", C_GREEN),
    ]

    def _kpi_cell(val, lbl, col):
        return Table(
            [[Paragraph(f'<font color="{col}"><b>{val}</b></font>',
                        ParagraphStyle("kv", fontName="Helvetica-Bold", fontSize=20,
                                       textColor=HexColor(col), alignment=1))],
             [Paragraph(lbl, ParagraphStyle("kl", fontName="Helvetica", fontSize=8,
                                            textColor=GRAY, alignment=1))]],
            colWidths=[CW / 4],
            style=TableStyle([
                ("TOPPADDING",    (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ("BOX",           (0, 0), (-1, -1), 0.8, LGRAY),
                ("BACKGROUND",    (0, 0), (-1, -1), OFFWHITE),
            ]))

    story.append(Table(
        [[_kpi_cell(v, l, c) for v, l, c in kpis]],
        colWidths=[CW / 4] * 4,
        style=TableStyle([
            ("LEFTPADDING",  (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ])))
    story.append(Spacer(1, 5 * mm))

    side = (CW - W_ALL) / 2
    story.append(Table(
        [[Spacer(1, 1), Image(donut_buf, width=W_ALL, height=H_ALL), Spacer(1, 1)]],
        colWidths=[side, W_ALL, side],
        style=TableStyle([
            ("ALIGN",         (1, 0), (1, 0),   "CENTER"),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING",    (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ])))
    story.append(Spacer(1, 3 * mm))
    story.append(Table(
        [[Image(barras_buf, width=W_ALL, height=H_ALL),
          Image(ucbar_buf,  width=W_ALL, height=H_ALL)]],
        colWidths=[CW / 2] * 2,
        style=TableStyle([
            ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING",   (0, 0), (-1, -1), 0),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
            ("TOPPADDING",    (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ])))
    story.append(PageBreak())

    # P4: Tabela por UC (pode ocupar múltiplas páginas com muitas UCs)
    story.append(Paragraph("Resultados por Unidade Consumidora", ST["section"]))
    story.append(HRFlowable(width="100%", thickness=1.5, color=LIME, spaceAfter=10))

    uc_header = ["UC", "Período", "Faturas", "OK", "DIVERGENTE"]
    uc_rows   = [uc_header]
    for uc, per, fat, ok, inv, div, rs in UCS:
        uc_rows.append([uc, per, str(fat), str(ok), str(inv + div)])
    uc_rows.append(["TOTAL", "—", str(N_FAT), str(TOT_OK), str(TOT_PROB)])

    row_styles = []
    for i in range(1, len(uc_rows) - 1):
        if i % 2 == 0:
            row_styles.append(("BACKGROUND", (0, i), (-1, i), HexColor("#F2F6FA")))

    uc_tbl = Table(uc_rows, colWidths=[CW / 5] * 5, repeatRows=1)
    uc_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0,  0), (-1,  0), NAVY),
        ("TEXTCOLOR",     (0,  0), (-1,  0), WHITE),
        ("FONTNAME",      (0,  0), (-1,  0), "Helvetica-Bold"),
        ("FONTSIZE",      (0,  0), (-1,  0), 9),
        ("ALIGN",         (0,  0), (-1,  0), "CENTER"),
        ("FONTNAME",      (0,  1), (-1, -2), "Helvetica"),
        ("FONTSIZE",      (0,  1), (-1, -2), 9),
        ("ALIGN",         (2,  1), (-1, -1), "CENTER"),
        ("ALIGN",         (0,  1), ( 1, -1), "LEFT"),
        ("BACKGROUND",    (0, -1), (-1, -1), NAVY),
        ("TEXTCOLOR",     (0, -1), (-1, -1), WHITE),
        ("FONTNAME",      (0, -1), (-1, -1), "Helvetica-Bold"),
        ("TEXTCOLOR",     (4,  1), ( 4, -2), BLUE),
        ("FONTNAME",      (4,  1), ( 4, -2), "Helvetica-Bold"),
        ("GRID",          (0,  0), (-1, -1), 0.5, LGRAY),
        ("TOPPADDING",    (0,  0), (-1, -1), 6),
        ("BOTTOMPADDING", (0,  0), (-1, -1), 6),
        ("LEFTPADDING",   (0,  0), (-1, -1), 6),
    ] + row_styles))
    story.append(uc_tbl)
    story.append(PageBreak())

    # P5: Andamento Processual
    story.append(Paragraph("Andamento Processual — Auditoria de Faturas", ST["section"]))
    story.append(HRFlowable(width="100%", thickness=1.5, color=LIME, spaceAfter=10))

    for i, (etapa, desc) in enumerate(PRAZOS, 1):
        bg = HexColor(C_NAVY if i % 2 == 1 else C_BLUE)
        cabec = Table(
            [[Paragraph(f'<font color="{C_LIME}"><b>Etapa {i}</b></font>',
                        ParagraphStyle("et", fontName="Helvetica-Bold",
                                       fontSize=9, textColor=LIME)),
              Paragraph(f'<font color="white">{etapa}</font>',
                        ParagraphStyle("en", fontName="Helvetica-Bold",
                                       fontSize=10, textColor=WHITE))]],
            colWidths=[50, CW - 50],
            style=TableStyle([
                ("BACKGROUND",    (0, 0), (-1, -1), bg),
                ("TOPPADDING",    (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("LEFTPADDING",   (0, 0), (-1, -1), 8),
            ]))
        corpo = Table(
            [[Paragraph(desc, ST["prazo_d"])]],
            colWidths=[CW],
            style=TableStyle([
                ("BACKGROUND",    (0, 0), (-1, -1), OFFWHITE),
                ("TOPPADDING",    (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ("LEFTPADDING",   (0, 0), (-1, -1), 10),
                ("BOX",           (0, 0), (-1, -1), 0.5, LGRAY),
            ]))
        story.append(cabec)
        story.append(corpo)
        story.append(Spacer(1, 2 * mm))

    story.append(Spacer(1, 4 * mm))
    story.append(Paragraph(OBS_PRAZOS, ST["obs"]))
    story.append(PageBreak())

    # Última página: Contracapa
    # _TriggerContracapa sinaliza ao canvas que esta página recebe a contracapa
    story.append(_TriggerContracapa())
    story.append(Spacer(1, 1))

    # Build
    buf = io.BytesIO()
    _cliente  = cliente_nome
    _dist     = DIST
    _periodo  = PERIODO
    _n_ucs    = N_UCS
    _data     = DATA_REL
    _parceiro = parceiro_nome.strip() if parceiro_nome else ""

    def _on_first_page(c, doc):
        _draw_capa(c, _cliente, _dist, _periodo, _n_ucs, _data, _parceiro)

    def _on_later_pages(c, doc):
        pass

    def _canvas_factory(filename, **kwargs):
        return _RelatorioCanvas(
            filename,
            cliente_nome=_cliente,
            parceiro_nome=_parceiro,
            **kwargs
        )

    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(A4),
        leftMargin=MARGIN,
        rightMargin=MARGIN,
        topMargin=26 * mm,
        bottomMargin=12 * mm,
        title=f"Relatório de Auditoria — {cliente_nome}",
        author="Minuto Energia",
    )
    doc.build(
        story,
        onFirstPage=_on_first_page,
        onLaterPages=_on_later_pages,
        canvasmaker=_canvas_factory,
    )

    return buf.getvalue()

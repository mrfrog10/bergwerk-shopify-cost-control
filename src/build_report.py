"""Build the monthly cost report (PDF + Markdown) and README charts from the
recalculated workbook. Every number and date is read through a named range,
so the report always matches the Excel model.

Run after verify_workbook.py:  python src/build_report.py
"""
import sys
from datetime import datetime
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from openpyxl import load_workbook  # noqa: E402
from reportlab.lib import colors  # noqa: E402
from reportlab.lib.enums import TA_LEFT  # noqa: E402
from reportlab.lib.pagesizes import A4  # noqa: E402
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet  # noqa
from reportlab.lib.units import mm  # noqa: E402
from reportlab.platypus import (Image, KeepTogether, PageBreak,  # noqa: E402
                                Paragraph, SimpleDocTemplate, Spacer, Table,
                                TableStyle)

sys.path.insert(0, str(Path(__file__).resolve().parent))
from verify_workbook import CALC, named, named_rows  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
REPORT_DIR, DOCS = ROOT / "report", ROOT / "docs"

GREEN, RED, AMBER, GREY = "#1F4E3D", "#C00000", "#F4B183", "#7F7F7F"

wb = load_workbook(CALC, data_only=True)
if named(wb, "INTEGRITY") != "ALL PASS":
    sys.exit("Model integrity checks fail: fix the workbook before reporting.")

K = {k: named(wb, k) for k in (
    "STATUS_DATE", "COST_RAG", "SCHED_RAG", "BAC", "PV", "EV", "AC", "CPI",
    "SPI", "EAC", "VAC", "RESERVE_LEFT", "CUR_TOTAL", "OVERRUN",
    "OVERRUN_PCT", "PLANNED_LAUNCH", "FORECAST_LAUNCH", "LAG_DAYS", "ES",
    "AT", "SPI_T", "SPI_WP5", "EAC_OPT", "EAC_CPI", "EAC_CPISPI", "TCPI_BAC",
    "TCPI_EAC", "COMMITTED", "PENDING_CR")}
STATUS = K["STATUS_DATE"]
SCHED = named_rows(wb, "SCHED_TABLE")
N_MONTHS = len(SCHED)
STAMP = STATUS.strftime("%Y-%m")
PDF = REPORT_DIR / f"BergWerk_Monthly_Cost_Report_{STAMP}.pdf"
MD = REPORT_DIR / f"BergWerk_Monthly_Cost_Report_{STAMP}.md"

WPS = []
for e, f in zip(named_rows(wb, "EV_TABLE"), named_rows(wb, "FC_TABLE")):
    WPS.append(dict(zip(
        ("wp", "name", "bac", "pct", "ev", "pv", "ac", "cv", "sv", "cpi",
         "spi"), e), etc=f[3], eac=f[5], vac=f[7], vacp=f[8]))
DIFFS = named_rows(wb, "DIFF_TABLE")
RISKS = named_rows(wb, "RISKS")
DECISIONS = named_rows(wb, "DECISIONS")
CHANGES = named_rows(wb, "CHANGELOG")
CHECKS = named_rows(wb, "CHECKS_TABLE")


def eur(v, signed=False):
    s = f"{abs(v):,.0f}"
    if signed:
        return ("+" if round(v) > 0 else "-" if round(v) < 0 else "") + s
    return ("-" if round(v) < 0 else "") + s


def d(v, fmt="%d %b %Y"):
    return v.strftime(fmt) if isinstance(v, datetime) else str(v)


def ratio(v):
    return v if isinstance(v, str) else f"{v:.2f}"


# ------------------------------------------------------------------ charts
DOCS.mkdir(exist_ok=True)
months = [r[0].strftime("%b") for r in SCHED]
cum_pv = [r[1] for r in SCHED]
cum_ac = [r[2] for r in SCHED if isinstance(r[2], (int, float))]
n_ac = len(cum_ac)

plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 9})
fig, ax = plt.subplots(figsize=(7.2, 3.4), dpi=200)
ax.plot(months, cum_pv, color=GREY, ls="--", lw=2, marker="o", ms=4,
        label="Planned value (PV)")
ax.plot(months[:n_ac], cum_ac, color=RED, lw=2.2, marker="o", ms=4,
        label="Actual cost (AC)")
ax.plot([months[n_ac - 1]], [K["EV"]], color=GREEN, marker="D", ms=9, ls="",
        label="Earned value (EV)")
top = max(K["EAC_CPISPI"], K["BAC"], max(cum_pv)) * 1.12
ax.axhline(K["BAC"], color=GREEN, lw=0.8, alpha=.5)
ax.text(0, K["BAC"] + top * .012, f"Budget {eur(K['BAC'])}", color=GREEN,
        fontsize=8)
ax.errorbar([N_MONTHS - 1], [K["EAC"]],
            yerr=[[K["EAC"] - K["EAC_OPT"]], [K["EAC_CPISPI"] - K["EAC"]]],
            fmt="x", color=RED, ms=9, mew=2, capsize=4, lw=1.2,
            label="Forecast final cost (range)")
ax.text(N_MONTHS - 1.15, K["EAC"], f"{eur(K['EAC'])}", color=RED, fontsize=8,
        ha="right", va="center")
ax.annotate(f"AC {eur(K['AC'])}", (n_ac - 1, K["AC"]), xytext=(8, 4),
            textcoords="offset points", color=RED, fontsize=8)
ax.annotate(f"EV {eur(K['EV'])}", (n_ac - 1, K["EV"]), xytext=(8, -12),
            textcoords="offset points", color=GREEN, fontsize=8)
ax.yaxis.set_major_formatter(lambda v, _: f"{v / 1000:.0f}k")
ax.set_ylim(0, top)
ax.set_ylabel("EUR, cumulative")
ax.spines[["top", "right"]].set_visible(False)
ax.grid(axis="y", alpha=.25)
ax.legend(frameon=False, loc="upper left", fontsize=7.5)
ax.set_title(f"Cumulative cost: plan vs actual vs value earned "
             f"(to {d(STATUS, '%d %b %Y')})", loc="left", fontsize=10,
             color=GREEN, weight="bold")
fig.tight_layout()
fig.savefig(DOCS / "s_curve.png")
plt.close(fig)

fig, ax = plt.subplots(figsize=(7.2, 3.0), dpi=200)
x = range(len(WPS))
ax.bar([i - .2 for i in x], [w["bac"] for w in WPS], .4, color="#A9C5B4",
       label="Budget")
ax.bar([i + .2 for i in x], [w["eac"] for w in WPS], .4,
       color=[RED if w["vac"] > 0 else GREEN for w in WPS],
       label="Expected final cost (red = over, green = within)")
for i, w in enumerate(WPS):
    if round(w["vac"]):
        ax.text(i + .2, w["eac"] + 400, f"{w['vacp']:+.0%}", ha="center",
                fontsize=8, color=RED if w["vac"] > 0 else GREEN,
                weight="bold")
ax.set_xticks(list(x), [w["wp"] for w in WPS])
ax.yaxis.set_major_formatter(lambda v, _: f"{v / 1000:.0f}k")
ax.spines[["top", "right"]].set_visible(False)
ax.grid(axis="y", alpha=.25)
ax.legend(frameon=False, fontsize=8)
ax.set_title("Budget vs expected final cost by work package", loc="left",
             fontsize=10, color=GREEN, weight="bold")
fig.tight_layout()
fig.savefig(DOCS / "wp_variance.png")
plt.close(fig)

# ------------------------------------------------------------- shared text
behind = (1 - K["SPI"]) * 100
SUMMARY = (
    f"Spending looks on plan: actual cost of {eur(K['AC'])} is "
    f"{eur(abs(K['AC'] - K['PV']))} {'below' if K['AC'] < K['PV'] else 'above'}"
    f" the planned {eur(K['PV'])}. Earned value shows a different picture. "
    f"Only {eur(K['EV'])} of planned work is done, so each euro spent has "
    f"delivered {K['CPI'] * 100:.0f} cents of work (CPI {K['CPI']:.2f}) and "
    f"the project is about {behind:.0f}% behind plan (SPI {K['SPI']:.2f}). "
    f"The bottom-up forecast is {eur(K['EAC'])} against a work-package "
    f"budget of {eur(K['BAC'])} ({eur(K['VAC'], True)}). After the remaining "
    f"reserve of {eur(K['RESERVE_LEFT'])} is spent, the project is still "
    f"{eur(K['OVERRUN'])} over its EUR {eur(K['CUR_TOTAL'])} budget.")
RANGE_NOTE = (
    f"The forecast is more likely to rise than fall. Reaching "
    f"{eur(K['EAC'])} needs a cost efficiency of {K['TCPI_EAC']:.2f} on the "
    f"remaining work, but only {K['CPI']:.2f} has been achieved so far. "
    f"Finishing within budget would need {K['TCPI_BAC']:.2f}, which is not "
    f"realistic. Other methods give {eur(K['EAC_OPT'])} to "
    f"{eur(K['EAC_CPISPI'])}.")
SCHED_NOTE = (
    f"Earned value equals the plan for month {K['ES']:.2f}, after "
    f"{K['AT']} months of work: a lag of {K['LAG_DAYS']} days. This moves "
    f"the launch from {d(K['PLANNED_LAUNCH'])} to about "
    f"{d(K['FORECAST_LAUNCH'])}. It assumes the lag does not grow. Testing "
    f"(WP5) is at SPI {K['SPI_WP5']:.2f}, so the new date is at risk.")
cr_bits = []
for cid, _, what, wp, amt, paid, status, phase, _note in CHANGES:
    src = "from the reserve" if paid == "Reserve" else f"paid by {paid}"
    if status == "Approved":
        cr_bits.append(f"{cid} ({what.split(' (')[0]}, EUR {eur(amt)} {src}, "
                       f"{wp}) is approved")
    else:
        cr_bits.append(f"{cid} ({what.split(' (')[0]}, EUR {eur(amt)} {src}) "
                       f"is {status.lower()} and is not in the baseline or "
                       f"the forecast")
CR_NOTE = "Change requests: " + "; ".join(cr_bits) + "."
COMMIT_NOTE = (f"Committed costs of {eur(K['COMMITTED'])} (ordered, not yet "
               f"invoiced) are included in the forecast, not added on top.")

# --------------------------------------------------------------------- PDF
REPORT_DIR.mkdir(exist_ok=True)
for old in REPORT_DIR.glob("BergWerk_Monthly_Cost_Report_*"):
    if old.stem != PDF.stem:
        old.unlink()
ss = getSampleStyleSheet()
H1 = ParagraphStyle("h1", parent=ss["Title"], fontSize=17, leading=21,
                    textColor=colors.HexColor(GREEN), alignment=TA_LEFT,
                    spaceAfter=2)
H2 = ParagraphStyle("h2", parent=ss["Heading2"], fontSize=11.5, leading=14,
                    textColor=colors.HexColor(GREEN), spaceBefore=8,
                    spaceAfter=4)
BODY = ParagraphStyle("b", parent=ss["BodyText"], fontSize=9, leading=12)
SMALL = ParagraphStyle("s", parent=BODY, fontSize=8, leading=10,
                       textColor=colors.HexColor("#595959"))
CELL = ParagraphStyle("c", parent=BODY, fontSize=8.2, leading=10)
CELLB = ParagraphStyle("cb", parent=CELL, fontName="Helvetica-Bold")
RAG = {"RED": colors.HexColor(RED), "AMBER": colors.HexColor(AMBER),
       "GREEN": colors.HexColor("#70AD47")}


def table(data, widths, right_cols=(), bold_last=False):
    right_cols = set(right_cols)
    rows = [[Paragraph(f"<font color='white'>{c}</font>", CELLB)
             for c in data[0]]]
    for i, row in enumerate(data[1:], start=1):
        last = bold_last and i == len(data) - 1
        rows.append([Paragraph(str(c), ParagraphStyle(
            f"t{i}{j}", parent=CELLB if last else CELL,
            alignment=2 if j in right_cols else 0))
            for j, c in enumerate(row)])
    st = [("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
          ("GRID", (0, 0), (-1, -1), .3, colors.HexColor("#BFBFBF")),
          ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(GREEN)),
          ("TOPPADDING", (0, 0), (-1, -1), 2.5),
          ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5)]
    st += [("BACKGROUND", (0, i), (-1, i), colors.HexColor("#F2F7F4"))
           for i in range(2, len(rows), 2)]
    if bold_last:
        st.append(("LINEABOVE", (0, -1), (-1, -1), 1, colors.black))
    return Table(rows, colWidths=widths, repeatRows=1, style=st)


def footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 7.5)
    canvas.setFillColor(colors.HexColor(GREY))
    canvas.drawString(18 * mm, 10 * mm,
                      f"BergWerk Outdoor GmbH (fictional client) | Shopify "
                      f"migration | Monthly cost report, status "
                      f"{d(STATUS)}")
    canvas.drawRightString(A4[0] - 18 * mm, 10 * mm, f"Page {doc.page}")
    canvas.restoreState()


doc = SimpleDocTemplate(str(PDF), pagesize=A4, leftMargin=18 * mm,
                        rightMargin=18 * mm, topMargin=15 * mm,
                        bottomMargin=17 * mm,
                        title=f"BergWerk Monthly Cost Report - "
                              f"{d(STATUS, '%B %Y')}",
                        author="Project Controls")
W = A4[0] - 36 * mm
story = [
    Paragraph("Monthly cost report: Shopify migration", H1),
    Paragraph(f"BergWerk Outdoor GmbH, Munich (fictional client) | Status "
              f"date {d(STATUS)} | Month {K['AT']} of {N_MONTHS} | All "
              f"figures in EUR", SMALL),
    Spacer(1, 6),
]
rag = Table([[Paragraph("<b>Cost</b>", CELL), K["COST_RAG"],
              Paragraph("<b>Schedule</b>", CELL), K["SCHED_RAG"],
              Paragraph(f"Forecast overrun after reserve <b>"
                        f"{eur(K['OVERRUN'], True)}</b> "
                        f"({K['OVERRUN_PCT']:.1%} of {eur(K['CUR_TOTAL'])})."
                        f" Launch likely <b>{d(K['FORECAST_LAUNCH'])}</b> "
                        f"instead of {d(K['PLANNED_LAUNCH'])}.", CELL)]],
            colWidths=[16 * mm, 20 * mm, 20 * mm, 20 * mm, W - 76 * mm])
rag.setStyle(TableStyle([
    ("BACKGROUND", (1, 0), (1, 0), RAG[K["COST_RAG"]]),
    ("BACKGROUND", (3, 0), (3, 0), RAG[K["SCHED_RAG"]]),
    ("TEXTCOLOR", (1, 0), (1, 0), colors.white),
    ("FONTNAME", (1, 0), (3, 0), "Helvetica-Bold"),
    ("ALIGN", (1, 0), (3, 0), "CENTER"),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("BOX", (0, 0), (-1, -1), .5, colors.HexColor("#BFBFBF")),
    ("TOPPADDING", (0, 0), (-1, -1), 5),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 5)]))
story += [rag, Paragraph("Summary", H2), Paragraph(SUMMARY, BODY)]

kp = [["Measure", "Value", "Meaning"],
      ["Budget at completion (BAC)", eur(K["BAC"]),
       "Six work packages after approved CRs; reserve held separately"],
      ["Planned value to date (PV)", eur(K["PV"]),
       f"Work planned by {d(STATUS)}"],
      ["Earned value (EV)", eur(K["EV"]), "% complete x budget"],
      ["Actual cost to date (AC)", eur(K["AC"]),
       "Internal hours at EUR 25 + external costs"],
      ["Cost so far vs planned (AC - PV)", eur(K["AC"] - K["PV"], True),
       "Looks fine on the surface"],
      ["CPI = EV / AC", f"{K['CPI']:.2f}",
       f"{K['CPI'] * 100:.0f} cents of planned work per euro"],
      ["SPI = EV / PV", f"{K['SPI']:.2f}", f"About {behind:.0f}% behind plan"],
      ["Expected final cost (EAC)", eur(K["EAC"]),
       f"Bottom-up; range {eur(K['EAC_OPT'])} to {eur(K['EAC_CPISPI'])}"],
      ["Variance at completion", eur(K["VAC"], True),
       f"Against {eur(K['BAC'])}"],
      ["Overrun after reserve", eur(K["OVERRUN"], True),
       f"{K['OVERRUN_PCT']:.1%} of the {eur(K['CUR_TOTAL'])} total"],
      ["TCPI to finish within budget", f"{K['TCPI_BAC']:.2f}",
       "Efficiency needed on remaining work; not realistic above 1.10"],
      ["Forecast launch", d(K["FORECAST_LAUNCH"]),
       f"{K['LAG_DAYS']} days after the planned {d(K['PLANNED_LAUNCH'])}"]]
story += [Paragraph("Key figures", H2),
          table(kp, [60 * mm, 24 * mm, W - 84 * mm], right_cols=(1,)),
          Spacer(1, 6),
          Image(str(DOCS / "s_curve.png"), width=W, height=W * 3.4 / 7.2),
          PageBreak()]

rows = [["Work package", "Budget", "% done", "EV", "PV", "AC", "CPI", "SPI",
         "EAC", "EAC - budget"]]
for w in WPS:
    rows.append([f"{w['wp']} {w['name']}", eur(w["bac"]), f"{w['pct']:.0%}",
                 eur(w["ev"]), eur(w["pv"]), eur(w["ac"]), ratio(w["cpi"]),
                 ratio(w["spi"]), eur(w["eac"]),
                 f"{eur(w['vac'], True)} ({w['vacp']:+.0%})"
                 if round(w["vac"]) else "0"])
rows.append(["Total", eur(K["BAC"]), f"{K['EV'] / K['BAC']:.0%}",
             eur(K["EV"]), eur(K["PV"]), eur(K["AC"]), f"{K['CPI']:.2f}",
             f"{K['SPI']:.2f}", eur(K["EAC"]), eur(K["VAC"], True)])
story += [Paragraph("Earned value by work package", H2),
          table(rows, [44 * mm] + [13.2 * mm] * 8 + [W - 44 * mm - 105.6 * mm],
                right_cols=range(1, 10), bold_last=True),
          Spacer(1, 4),
          Image(str(DOCS / "wp_variance.png"), width=W * .82,
                height=W * .82 * 3.0 / 7.2)]

story += [Paragraph("Main differences from budget", H2),
          table([["Work package", "EAC - budget", "Reason"]] +
                [[a, eur(b, True), c] for a, b, c in DIFFS if round(b)],
                [55 * mm, 22 * mm, W - 77 * mm], right_cols=(1,))]

fr = [["Forecast method", "EAC", "Overrun after reserve"],
      ["Optimistic: remaining work costs as planned", eur(K["EAC_OPT"]),
       eur(K["EAC_OPT"] - K["BAC"] - K["RESERVE_LEFT"], True)],
      ["Bottom-up team estimates (reported)", eur(K["EAC"]),
       eur(K["OVERRUN"], True)],
      ["Cost efficiency so far continues (BAC / CPI)", eur(K["EAC_CPI"]),
       eur(K["EAC_CPI"] - K["BAC"] - K["RESERVE_LEFT"], True)],
      ["Cost and schedule pressure continue", eur(K["EAC_CPISPI"]),
       eur(K["EAC_CPISPI"] - K["BAC"] - K["RESERVE_LEFT"], True)]]
story += [KeepTogether([Paragraph("How reliable is the forecast?", H2),
                        table(fr, [90 * mm, 30 * mm, W - 120 * mm],
                              right_cols=(1, 2)),
                        Spacer(1, 3), Paragraph(RANGE_NOTE, BODY),
                        Spacer(1, 3), Paragraph(SCHED_NOTE, BODY)])]

story += [PageBreak(), Paragraph("Top 3 risks and one watch item", H2),
          table([["Risk", "Impact", "Likelihood", "Response"]] +
                [list(r) for r in RISKS],
                [66 * mm, 16 * mm, 19 * mm, W - 101 * mm])]
drows = [["Decision", "Effect", "Proposed due", "Recommendation"]]
for dec, eff, due, rec in DECISIONS:
    drows.append([dec, d(eff) if isinstance(eff, datetime) else
                  f"EUR {eur(eff)}", d(due, "%d %b"), rec])
story += [Paragraph("Decisions needed from the project manager", H2),
          table(drows, [56 * mm, 22 * mm, 18 * mm, W - 96 * mm]),
          Spacer(1, 6), Paragraph(CR_NOTE, SMALL), Spacer(1, 2),
          Paragraph(COMMIT_NOTE, SMALL)]

story += [Paragraph("Check figures (model vs project brief)", H2)]
crow = [["Measure", "Model", "Expected", "Result"]]
for m, _, a, b, res in CHECKS:
    def fmt(v):
        if isinstance(v, datetime):
            return d(v)
        if isinstance(v, float) and abs(v) < 2:
            return f"{v:.3g}"
        return eur(v, True) if isinstance(v, (int, float)) else str(v)
    crow.append([m, fmt(a), fmt(b), res])
story.append(table(crow, [70 * mm, 30 * mm, 30 * mm, W - 130 * mm],
                   right_cols=(1, 2)))
doc.build(story, onFirstPage=footer, onLaterPages=footer)

# ---------------------------------------------------------------- Markdown
emoji = {"RED": "🔴", "AMBER": "🟠", "GREEN": "🟢"}
md = [
    "# Monthly cost report: BergWerk Outdoor Shopify migration", "",
    f"**Status date:** {d(STATUS)} (month {K['AT']} of {N_MONTHS}) · "
    "**Client:** BergWerk Outdoor GmbH, Munich (fictional) · all figures "
    "in EUR", "",
    "| Cost | Schedule |", "|---|---|",
    f"| {emoji[K['COST_RAG']]} **{K['COST_RAG']}** | "
    f"{emoji[K['SCHED_RAG']]} **{K['SCHED_RAG']}** |", "",
    "## Summary", "", SUMMARY, "",
    "![S-curve](../docs/s_curve.png)", "",
    "## Earned value by work package", "",
    "| WP | Budget | % done | EV | PV | AC | CPI | SPI | EAC | EAC − budget |",
    "|---|--:|--:|--:|--:|--:|--:|--:|--:|--:|",
]
for w in WPS:
    vac = (f"{eur(w['vac'], True)} ({w['vacp']:+.0%})" if round(w["vac"])
           else "0")
    md.append(f"| {w['wp']} {w['name']} | {eur(w['bac'])} | {w['pct']:.0%} "
              f"| {eur(w['ev'])} | {eur(w['pv'])} | {eur(w['ac'])} | "
              f"{ratio(w['cpi'])} | {ratio(w['spi'])} | {eur(w['eac'])} | "
              f"{vac} |")
md += [f"| **Total** | **{eur(K['BAC'])}** | | **{eur(K['EV'])}** | "
       f"**{eur(K['PV'])}** | **{eur(K['AC'])}** | **{K['CPI']:.2f}** | "
       f"**{K['SPI']:.2f}** | **{eur(K['EAC'])}** | "
       f"**{eur(K['VAC'], True)}** |", "",
       "![WP variance](../docs/wp_variance.png)", "",
       "## Main differences from budget", ""]
md += [f"- **{a}, {eur(b, True)}:** {c}." for a, b, c in DIFFS if round(b)]
md += ["", "## How reliable is the forecast?", "",
       "| Method | EAC | Overrun after reserve |", "|---|--:|--:|"]
md += [f"| {r[0]} | {r[1]} | {r[2]} |" for r in fr[1:]]
md += ["", RANGE_NOTE, "", SCHED_NOTE, "", "## Top 3 risks and one watch "
       "item", ""]
md += [f"{i}. {r[0]} (impact {r[1].lower()}, likelihood {r[2].lower()}). "
       f"*Response:* {r[3]}." for i, r in enumerate(RISKS, 1)]
md += ["", "## Decisions needed from the project manager", ""]
for i, (dec, eff, due, rec) in enumerate(DECISIONS, 1):
    e = d(eff) if isinstance(eff, datetime) else f"EUR {eur(eff)}"
    md.append(f"{i}. **{dec}** ({e}; proposed due {d(due, '%d %b')}). "
              f"*Recommendation:* {rec}.")
md += ["", CR_NOTE, "", COMMIT_NOTE, "", "## Check figures", "",
       "| Measure | Model | Expected | Result |", "|---|--:|--:|:-:|"]
md += [f"| {r[0]} | {r[1]} | {r[2]} | {r[3]} |" for r in crow[1:]]
MD.write_text("\n".join(md) + "\n", encoding="utf-8")
print(f"Saved {PDF}\nSaved {MD}")

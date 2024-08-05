"""Build the BergWerk Outdoor Shopify migration cost-control workbook.

Every calculated figure is a live Excel formula, so the model updates when
an input (blue cell) changes. Key results are exposed as named ranges
(CPI, EAC, FORECAST_LAUNCH, ...) so the report never depends on cell
positions.

Run:  python src/build_workbook.py
"""
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.marker import Marker
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(Path(__file__).resolve().parent))
from inputs import BASE  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "excel" / "BergWerk_Shopify_Cost_Control.xlsx"

# ----------------------------------------------------------------- styles
GREEN_DARK = "1F4E3D"
GREEN_LIGHT = "E3EFE8"
INPUT_FILL = "FFF8DC"
TITLE = Font(name="Calibri", size=16, bold=True, color=GREEN_DARK)
SUB = Font(name="Calibri", size=10, italic=True, color="595959")
HDR_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor=GREEN_DARK)
SEC_FONT = Font(name="Calibri", size=11, bold=True, color=GREEN_DARK)
INPUT_FONT = Font(name="Calibri", size=10, color="0000FF")
CALC_FONT = Font(name="Calibri", size=10, color="000000")
HELPER_FONT = Font(name="Calibri", size=8, color="A6A6A6")
BOLD = Font(name="Calibri", size=10, bold=True)
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
TOP = Border(top=Side(style="thin", color="000000"),
             bottom=Side(style="double", color="000000"))
EUR = '#,##0;[Red]-#,##0;"-"'
EUR_SIGNED = '+#,##0;[Red]-#,##0;0'
PCT = '0%'
PCT1 = '0.0%'
RATIO = '0.00'
DATE_FMT = 'dd mmm yyyy'
MONTH_FMT = 'mmm yyyy'
OK_FILL = PatternFill("solid", fgColor="C6EFCE")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")


def title(ws, text, subtitle):
    ws["A1"] = text
    ws["A1"].font = TITLE
    ws["A2"] = subtitle
    ws["A2"].font = SUB
    ws.sheet_view.showGridLines = False


def header(ws, row, labels, col=1):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=lab)
        c.font, c.fill, c.border = HDR_FONT, HDR_FILL, BOX
        c.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
    ws.row_dimensions[row].height = 30


def cell(ws, ref, value, fmt=None, inp=False, bold=False, fill=None,
         border=True, align=None, font=None):
    c = ws[ref]
    c.value = value
    c.font = font or (INPUT_FONT if inp else (BOLD if bold else CALC_FONT))
    if inp:
        c.fill = PatternFill("solid", fgColor=INPUT_FILL)
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    if fmt:
        c.number_format = fmt
    if border:
        c.border = BOX
    if align:
        c.alignment = Alignment(horizontal=align, vertical="center",
                                wrap_text=True)
    return c


def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w


def total_row(ws, row, first_col, last_col, top, bottom, label_col="B",
              label="Total", fmt=EUR):
    ws[f"{label_col}{row}"] = label
    ws[f"{label_col}{row}"].font = BOLD
    for ci in range(first_col, last_col + 1):
        L = get_column_letter(ci)
        c = ws[f"{L}{row}"]
        c.value, c.font, c.number_format, c.border = (
            f"=SUM({L}{top}:{L}{bottom})", BOLD, fmt, TOP)


def ok_format(ws, rng, first_cell):
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'OR({first_cell}="OK",ISNUMBER(SEARCH("PASS",{first_cell})))'],
        fill=OK_FILL, font=Font(color="006100", bold=True)))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'OR({first_cell}="CHECK",{first_cell}="FAIL",'
                 f'{first_cell}="REVIEW",{first_cell}="OVERDRAWN",'
                 f'{first_cell}="Add phasing month")'],
        fill=BAD_FILL, font=Font(color="9C0006", bold=True)))


def build(s=BASE, out=OUT):
    wb = Workbook()
    names = {}

    def name(n, sheet, ref):
        names[n] = f"{sheet}!{ref}"

    nwp = len(s["wps"])
    months = s["months"]
    MCOLS = [get_column_letter(c) for c in range(3, 3 + len(months))]
    LASTM = MCOLS[-1]
    W1, WN = 5, 4 + nwp  # first/last WP row on WP-table sheets
    TR = WN + 1          # total row

    # ============================================================ Settings
    ws = wb.active
    ws.title = "Settings"
    title(ws, "Project settings",
          "Blue cells on a cream background are inputs. Every other number "
          "in the workbook is a formula.")
    rows = [
        ("Client", s["client"], None),
        ("Project", s["project"], None),
        ("Markets", s["markets"], None),
        ("Project start", s["start"], DATE_FMT),
        ("Project end", s["end"], DATE_FMT),
        ("Planned launch (go-live)", s["launch"], DATE_FMT),
        ("Status date (report date)", s["status"], DATE_FMT),
        ("Internal hourly rate (EUR)", s["hourly_rate"], EUR),
        ("Original contract budget incl. reserve (EUR)",
         s["contract_budget"], EUR),
        ("Schedule GREEN if SPI at least", s["spi_green"], RATIO),
        ("Schedule AMBER if SPI at least", s["spi_amber"], RATIO),
        ("Aftercare after launch (weeks)", s["aftercare_weeks"], '0'),
    ]
    header(ws, 3, ["Setting", "Value"])
    for i, (k, v, f) in enumerate(rows, start=4):
        cell(ws, f"A{i}", k)
        cell(ws, f"B{i}", v, fmt=f, inp=True, align="left")
    ws["A16"] = "RAG rules"
    ws["A16"].font = SEC_FONT
    for i, t in enumerate([
        "Cost RED: forecast overrun is larger than the remaining reserve",
        "Cost AMBER: forecast is over the work-package budget but the "
        "reserve covers it",
        "Cost GREEN: forecast within the work-package budget",
        "Schedule: by SPI against the two thresholds above",
    ], start=17):
        ws[f"A{i}"] = t
        ws[f"A{i}"].font = CALC_FONT
    ws["A22"] = "Colour key"
    ws["A22"].font = SEC_FONT
    cell(ws, "A23", "Input (change me)", inp=True)
    cell(ws, "A24", "Formula (do not overwrite)")
    ws["A26"] = ("Cost basis: internal hours at EUR 25/hour plus external "
                 "costs (app licences, freelancers, translation).")
    ws["A26"].font = SUB
    widths(ws, {"A": 52, "B": 75})
    LAUNCH, STATUS = "Settings!$B$9", "Settings!$B$10"
    CONTRACT, SPI_G, SPI_A = ("Settings!$B$12", "Settings!$B$13",
                              "Settings!$B$14")
    name("STATUS_DATE", "Settings", "$B$10")
    name("PLANNED_LAUNCH", "Settings", "$B$9")
    PLANNED_END, AFTERCARE = "Settings!$B$8", "Settings!$B$15"
    name("PLANNED_END", "Settings", "$B$8")

    # =========================================================== ChangeLog
    cl = wb.create_sheet("ChangeLog")
    title(cl, "Change log",
          "Approved changes move budget into work packages and need a "
          "phasing month for the revised baseline. Pending changes are not "
          "in the baseline or the forecast.")
    header(cl, 4, ["ID", "Date", "Change", "WP", "Amount (EUR)", "Paid by",
                   "Status", "Phasing month", "Note", "Check"])
    approved_months = {cr: months[mi] for cr, (wi, mi)
                       in s["cr_phasing"].items()}
    for i, ch in enumerate(s["changes"], start=5):
        row = list(ch[:7]) + [approved_months.get(ch[0]), ch[7]]
        for j, v in enumerate(row):
            L = get_column_letter(j + 1)
            cell(cl, f"{L}{i}", v, inp=True,
                 fmt={1: DATE_FMT, 4: EUR, 7: MONTH_FMT}.get(j),
                 align="left" if j in (2, 8) else None)
        cl.row_dimensions[i].height = 32
    for i in range(5, 31):
        cl[f"J{i}"] = (f'=IF(A{i}="","",IF(AND(G{i}="Approved",H{i}=""),'
                       f'"Add phasing month","OK"))')
        cl[f"J{i}"].alignment = Alignment(horizontal="center")
    ok_format(cl, "J5:J30", "J5")
    for rng, opts in (("G5:G30", '"Approved,Pending,Rejected"'),
                      ("F5:F30", '"Reserve,BergWerk"')):
        dv = DataValidation(type="list", formula1=opts, allow_blank=True)
        cl.add_data_validation(dv)
        dv.add(rng)
    widths(cl, {"A": 8, "B": 13, "C": 46, "D": 7, "E": 13, "F": 11,
                "G": 11, "H": 13, "I": 46, "J": 17})
    CL_ID, CL_WP, CL_AMT = ("ChangeLog!$A$5:$A$30", "ChangeLog!$D$5:$D$30",
                            "ChangeLog!$E$5:$E$30")
    CL_PAID, CL_STATUS, CL_MONTH = ("ChangeLog!$F$5:$F$30",
                                    "ChangeLog!$G$5:$G$30",
                                    "ChangeLog!$H$5:$H$30")
    name("CHANGELOG", "ChangeLog", f"$A$5:$I${4 + len(s['changes'])}")

    # ============================================================== Budget
    bd = wb.create_sheet("Budget", 1)
    title(bd, "Budget by work package",
          "Original plan plus approved change requests. Reserve and contract "
          "value reconciliation below.")
    header(bd, 4, ["WP", "Work package", "What it covers", "Months",
                   "Original budget", "Approved CRs", "Current budget"])
    for i, (wp, nm, scope, mo, budget) in enumerate(s["wps"], start=W1):
        cell(bd, f"A{i}", wp, inp=True, align="center")
        cell(bd, f"B{i}", nm, inp=True)
        cell(bd, f"C{i}", scope, inp=True, align="left")
        cell(bd, f"D{i}", mo, inp=True, align="center")
        cell(bd, f"E{i}", budget, EUR, inp=True)
        cell(bd, f"F{i}",
             f'=SUMIFS({CL_AMT},{CL_WP},A{i},{CL_STATUS},"Approved")', EUR)
        cell(bd, f"G{i}", f"=E{i}+F{i}", EUR)
        bd.row_dimensions[i].height = 30
    total_row(bd, TR, 5, 7, W1, WN, label="Work packages total")
    b = TR + 2  # 13
    cell(bd, f"B{b}", "Reserve for surprises (original)")
    cell(bd, f"E{b}", s["reserve"], EUR, inp=True)
    cell(bd, f"B{b + 1}", "Reserve used by approved CRs")
    cell(bd, f"E{b + 1}",
         f'=SUMIFS({CL_AMT},{CL_PAID},"Reserve",{CL_STATUS},"Approved")', EUR)
    cell(bd, f"B{b + 2}", "Reserve remaining", bold=True)
    cell(bd, f"E{b + 2}", f"=E{b}-E{b + 1}", EUR, bold=True)
    cell(bd, f"B{b + 3}", "Reserve check")
    cell(bd, f"E{b + 3}", f'=IF(E{b + 2}<0,"OVERDRAWN","OK")', align="center")
    cell(bd, f"B{b + 5}", "Client-funded approved CRs (raise contract value)")
    cell(bd, f"E{b + 5}",
         f'=SUMIFS({CL_AMT},{CL_PAID},"<>Reserve",{CL_STATUS},"Approved")',
         EUR)
    cell(bd, f"B{b + 6}", "Pending CRs (not in baseline or forecast)")
    cell(bd, f"E{b + 6}", f'=SUMIFS({CL_AMT},{CL_STATUS},"Pending")', EUR)
    t = b + 8  # 21
    for L, lab in (("E", "Original"), ("G", "Current")):
        bd[f"{L}{t - 1}"] = lab
        bd[f"{L}{t - 1}"].font = BOLD
        bd[f"{L}{t - 1}"].alignment = Alignment(horizontal="right")
    cell(bd, f"B{t}", "Work packages + reserve", bold=True)
    cell(bd, f"E{t}", f"=E{TR}+E{b}", EUR, bold=True)
    cell(bd, f"G{t}", f"=G{TR}+E{b + 2}", EUR, bold=True)
    cell(bd, f"B{t + 1}", "Contract value (original + client-funded CRs)",
         bold=True)
    cell(bd, f"E{t + 1}", f"={CONTRACT}", EUR, bold=True)
    cell(bd, f"G{t + 1}", f"={CONTRACT}+E{b + 5}", EUR, bold=True)
    cell(bd, f"B{t + 2}", "Check: budget reconciles to contract value")
    cell(bd, f"E{t + 2}", f'=IF(E{t}=E{t + 1},"OK","CHECK")', align="center")
    cell(bd, f"G{t + 2}", f'=IF(G{t}=G{t + 1},"OK","CHECK")', align="center")
    ok_format(bd, f"E{b + 3}:G{t + 2}", f"E{b + 3}")
    widths(bd, {"A": 7, "B": 46, "C": 58, "D": 10, "E": 15, "F": 14,
                "G": 15})
    RESERVE_LEFT = f"Budget!$E${b + 2}"
    CUR_TOTAL = f"Budget!$G${t + 1}"
    name("RESERVE_LEFT", "Budget", f"$E${b + 2}")
    name("CUR_TOTAL", "Budget", f"$G${t + 1}")
    name("PENDING_CR", "Budget", f"$E${b + 6}")
    BUDGET_CHECKS = (f"Budget!$E${b + 3}", f"Budget!$E${t + 2}",
                     f"Budget!$G${t + 2}")

    # ======================================================== PlannedValue
    pv = wb.create_sheet("PlannedValue", 2)
    title(pv, "Planned cost per month (Planned Value, PV)",
          "Original baseline + phasing of approved change requests (from the "
          "change log) = revised baseline used for earned value.")
    TOT = get_column_letter(3 + len(months))       # Total col
    TODATE = get_column_letter(4 + len(months))    # PV to date
    CHK = get_column_letter(5 + len(months))

    def block(top, label, kind, hdr_row=None):
        pv[f"A{top}"] = label
        pv[f"A{top}"].font = SEC_FONT
        h = top + 1
        header(pv, h, ["WP", "Work package"] + [""] * len(months) + ["Total"])
        for k, L in enumerate(MCOLS):
            c = pv[f"{L}{h}"]
            c.value = months[k] if hdr_row is None else f"={L}${hdr_row}"
            c.number_format = MONTH_FMT
        first = h + 1
        for i in range(nwp):
            r = first + i
            cell(pv, f"A{r}", f"=Budget!A{W1 + i}", align="center")
            cell(pv, f"B{r}", f"=Budget!B{W1 + i}")
            for k, L in enumerate(MCOLS):
                if kind == "input":
                    cell(pv, f"{L}{r}", s["pv"][i][k], EUR, inp=True)
                elif kind == "cr":
                    cell(pv, f"{L}{r}",
                         f'=SUMIFS({CL_AMT},{CL_WP},$A{r},{CL_STATUS},'
                         f'"Approved",{CL_MONTH},{L}${h})', EUR)
                else:
                    cell(pv, f"{L}{r}", f"={L}{kind[0] + i}+{L}{kind[1] + i}",
                         EUR)
            cell(pv, f"{TOT}{r}", f"=SUM(C{r}:{LASTM}{r})", EUR, bold=True)
        last = first + nwp - 1
        total_row(pv, last + 1, 3, 3 + len(months), first, last)
        return h, first, last

    h1, o1, _ = block(4, "A. Original baseline", "input")
    h2, c1, _ = block(h1 + nwp + 3, "B. Phasing of approved change requests",
                      "cr", hdr_row=h1)
    h3, r1, r6 = block(h2 + nwp + 3, "C. Revised baseline (used for EVM)",
                       (o1, c1), hdr_row=h1)
    for L, lab in ((TODATE, "PV to status date"),
                   (CHK, "Check vs current budget")):
        c = pv[f"{L}{h3}"]
        c.value, c.font, c.fill, c.border = lab, HDR_FONT, HDR_FILL, BOX
        c.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
    for i in range(nwp):
        r = r1 + i
        cell(pv, f"{TODATE}{r}",
             f'=SUMIFS(C{r}:{LASTM}{r},$C${h3}:${LASTM}${h3},"<="&{STATUS})',
             EUR, bold=True)
        cell(pv, f"{CHK}{r}", f'=IF({TOT}{r}=Budget!G{W1 + i},"OK","CHECK")',
             align="center")
    ok_format(pv, f"{CHK}{r1}:{CHK}{r6}", f"{CHK}{r1}")
    rt = r6 + 1
    c = pv[f"{TODATE}{rt}"]
    c.value, c.font, c.number_format, c.border = (
        f"=SUM({TODATE}{r1}:{TODATE}{r6})", BOLD, EUR, TOP)
    cum = rt + 1
    pv[f"B{cum}"] = "Cumulative PV"
    pv[f"B{cum}"].font = BOLD
    for k, L in enumerate(MCOLS):
        prev = f"+{MCOLS[k - 1]}{cum}" if k else ""
        cell(pv, f"{L}{cum}", f"={L}{rt}{prev}", EUR, bold=True,
             fill=GREEN_LIGHT)
    widths(pv, {"A": 7, "B": 32, **{L: 12 for L in MCOLS}, TOT: 12,
                TODATE: 14, CHK: 13})
    PV_TODATE = [f"PlannedValue!${TODATE}${r1 + i}" for i in range(nwp)]
    PV_CHECK_RNG = f"PlannedValue!${CHK}${r1}:${CHK}${r6}"

    # ========================================================== ActualCost
    ac = wb.create_sheet("ActualCost", 3)
    title(ac, "Actual costs (Actual Cost, AC)",
          "Internal hours at EUR 25/hour plus invoiced external costs. Only "
          "months up to the status date count towards AC.")
    header(ac, 4, ["WP", "Work package"] + [""] * len(months) +
           ["Total booked", "AC to status date"])
    for k, L in enumerate(MCOLS):
        ac[f"{L}4"] = f"=PlannedValue!{L}${h1}"
        ac[f"{L}4"].number_format = MONTH_FMT
    for i in range(nwp):
        r = W1 + i
        cell(ac, f"A{r}", f"=Budget!A{W1 + i}", align="center")
        cell(ac, f"B{r}", f"=Budget!B{W1 + i}")
        for k, L in enumerate(MCOLS):
            cell(ac, f"{L}{r}", s["ac"][i][k], EUR, inp=True)
        cell(ac, f"{TOT}{r}", f"=SUM(C{r}:{LASTM}{r})", EUR, bold=True)
        cell(ac, f"{TODATE}{r}",
             f'=SUMIFS(C{r}:{LASTM}{r},$C$4:${LASTM}$4,"<="&{STATUS})', EUR,
             bold=True)
    total_row(ac, TR, 3, 4 + len(months), W1, WN)
    ac[f"B{TR + 1}"] = "Cumulative AC"
    ac[f"B{TR + 1}"].font = BOLD
    for k, L in enumerate(MCOLS):
        prev = f"+{MCOLS[k - 1]}{TR + 1}" if k else ""
        cell(ac, f"{L}{TR + 1}", f"=IF({L}$4<={STATUS},{L}{TR}{prev},NA())",
             EUR, bold=True, fill=GREEN_LIGHT)
    n0 = TR + 3
    ac[f"A{n0}"] = "What happened"
    ac[f"A{n0}"].font = SEC_FONT
    header(ac, n0 + 1, ["Month", "", "Explanation"])
    for k, note in enumerate(s["ac_notes"]):
        r = n0 + 2 + k
        cell(ac, f"A{r}", f"={MCOLS[k]}4", MONTH_FMT)
        ac.merge_cells(f"A{r}:B{r}")
        cell(ac, f"C{r}", note, inp=True, align="left")
        ac.merge_cells(f"C{r}:{TODATE}{r}")
        ac.row_dimensions[r].height = 30
    widths(ac, {"A": 7, "B": 32, **{L: 12 for L in MCOLS}, TOT: 12,
                TODATE: 14})
    AC_TODATE = [f"ActualCost!${TODATE}${W1 + i}" for i in range(nwp)]
    AC_FUTURE_OK = f"ActualCost!${TOT}${TR}=ActualCost!${TODATE}${TR}"

    # ========================================================= Commitments
    cm = wb.create_sheet("Commitments", 5)
    title(cm, "Committed costs at status date",
          "Ordered but not yet invoiced. Included in the estimate to "
          "complete on the Forecast sheet, so never added on top of it.")
    header(cm, 4, ["Item", "WP", "Amount (EUR)"])
    nc = len(s["commitments"])
    for i, (item, wp, amt) in enumerate(s["commitments"], start=5):
        cell(cm, f"A{i}", item, inp=True)
        cell(cm, f"B{i}", wp, inp=True, align="center")
        cell(cm, f"C{i}", amt, EUR, inp=True)
    total_row(cm, 5 + nc, 3, 3, 5, 4 + nc, label_col="A")
    cm["A" + str(7 + nc)] = ("Note: the Shopify app licence is a yearly "
                             "fee; it is charged to the project because it "
                             "is in the agreed WP3 scope.")
    cm["A" + str(7 + nc)].font = SUB
    widths(cm, {"A": 56, "B": 8, "C": 15})
    COMMIT_TOTAL = f"Commitments!$C${5 + nc}"

    # ========================================================= EarnedValue
    ev = wb.create_sheet("EarnedValue", 6)
    title(ev, "Earned value at status date",
          "EV = % complete x current budget. CV = EV - AC, SV = EV - PV, "
          "CPI = EV / AC, SPI = EV / PV.")
    header(ev, 4, ["WP", "Work package", "Current budget (BAC)",
                   "% complete", "Earned value (EV)", "Planned value (PV)",
                   "Actual cost (AC)", "Cost variance (CV)",
                   "Schedule variance (SV)", "CPI", "SPI"])
    for i in range(nwp):
        r = W1 + i
        cell(ev, f"A{r}", f"=Budget!A{W1 + i}", align="center")
        cell(ev, f"B{r}", f"=Budget!B{W1 + i}")
        cell(ev, f"C{r}", f"=Budget!G{W1 + i}", EUR)
        cell(ev, f"D{r}", s["pct_complete"][i], PCT, inp=True)
        cell(ev, f"E{r}", f"=C{r}*D{r}", EUR, bold=True)
        cell(ev, f"F{r}", f"={PV_TODATE[i]}", EUR)
        cell(ev, f"G{r}", f"={AC_TODATE[i]}", EUR)
        cell(ev, f"H{r}", f"=E{r}-G{r}", EUR_SIGNED)
        cell(ev, f"I{r}", f"=E{r}-F{r}", EUR_SIGNED)
        cell(ev, f"J{r}", f'=IF(G{r}=0,"n/a",E{r}/G{r})', RATIO,
             align="center")
        cell(ev, f"K{r}", f'=IF(F{r}=0,"n/a",E{r}/F{r})', RATIO,
             align="center")
    dv = DataValidation(type="decimal", operator="between", formula1="0",
                        formula2="1", showErrorMessage=True,
                        errorTitle="% complete",
                        error="Enter a value between 0% and 100%.")
    ev.add_data_validation(dv)
    dv.add(f"D{W1}:D{WN}")
    total_row(ev, TR, 3, 9, W1, WN)
    ev[f"D{TR}"] = f"=E{TR}/C{TR}"
    ev[f"D{TR}"].number_format = PCT
    for L, f in (("J", f"=E{TR}/G{TR}"), ("K", f"=E{TR}/F{TR}")):
        c = ev[f"{L}{TR}"]
        c.value, c.font, c.number_format, c.border = f, BOLD, RATIO, TOP
        c.alignment = Alignment(horizontal="center")
    ev[f"B{TR + 2}"] = ("Reading: CPI below 1.00 = each euro spent buys less "
                        "than a euro of planned work. SPI below 1.00 = less "
                        "work done than planned.")
    ev[f"B{TR + 2}"].font = SUB
    ev.conditional_formatting.add(f"J{W1}:K{TR}", CellIsRule(
        operator="lessThan", formula=["0.9"],
        fill=PatternFill("solid", fgColor="F8CBAD")))
    ev.conditional_formatting.add(f"J{W1}:K{TR}", CellIsRule(
        operator="between", formula=["0.9", "0.9999"],
        fill=PatternFill("solid", fgColor="FFE699")))
    widths(ev, {"A": 7, "B": 32, "C": 14, "D": 11, "E": 14, "F": 14,
                "G": 14, "H": 14, "I": 14, "J": 8, "K": 8})
    for n, col in (("BAC", "C"), ("EV", "E"), ("PV", "F"), ("AC", "G"),
                   ("CPI", "J"), ("SPI", "K")):
        name(n, "EarnedValue", f"${col}${TR}")
    name("EV_TABLE", "EarnedValue", f"$A${W1}:$K${WN}")
    EV_T, PV_T, AC_T = (f"EarnedValue!$E${TR}", f"EarnedValue!$F${TR}",
                        f"EarnedValue!$G${TR}")
    BAC_T = f"EarnedValue!$C${TR}"
    CPI, SPI = f"EarnedValue!$J${TR}", f"EarnedValue!$K${TR}"
    PCT_RNG = f"EarnedValue!$D${W1}:$D${WN}"

    # ============================================================ Forecast
    fc = wb.create_sheet("Forecast", 7)
    title(fc, "Forecast of the final cost (Estimate at Completion, EAC)",
          "EAC = actual cost so far + estimate to complete (ETC). ETC "
          "already includes committed costs.")
    header(fc, 4, ["WP", "Work package", "Spent so far (AC)",
                   "Still to come (ETC)", "of which committed",
                   "Expected final (EAC)", "Current budget (BAC)",
                   "Difference (VAC)", "Difference %",
                   "ETC covers committed?", "Explanation", "Rank key"])
    for i in range(nwp):
        r = W1 + i
        cell(fc, f"A{r}", f"=Budget!A{W1 + i}", align="center")
        cell(fc, f"B{r}", f"=Budget!B{W1 + i}")
        cell(fc, f"C{r}", f"={AC_TODATE[i]}", EUR)
        cell(fc, f"D{r}", s["etc"][i], EUR, inp=True)
        cell(fc, f"E{r}", f"=SUMIFS(Commitments!$C$5:$C$40,"
             f"Commitments!$B$5:$B$40,A{r})", EUR)
        cell(fc, f"F{r}", f"=C{r}+D{r}", EUR, bold=True)
        cell(fc, f"G{r}", f"=Budget!G{W1 + i}", EUR)
        cell(fc, f"H{r}", f"=F{r}-G{r}", EUR_SIGNED, bold=True)
        cell(fc, f"I{r}", f"=IF(G{r}=0,0,H{r}/G{r})", '+0%;[Red]-0%;0%')
        cell(fc, f"J{r}", f'=IF(D{r}>=E{r},"OK","CHECK")', align="center")
        cell(fc, f"K{r}", s["explanations"][i], inp=True, align="left")
        cell(fc, f"L{r}", f"=H{r}-ROW()/1000000", '0.000000',
             font=HELPER_FONT)
        fc.row_dimensions[r].height = 30
    ok_format(fc, f"J{W1}:J{WN}", f"J{W1}")
    total_row(fc, TR, 3, 8, W1, WN)
    fc[f"I{TR}"] = f"=H{TR}/G{TR}"
    fc[f"I{TR}"].number_format = '+0.0%;[Red]-0.0%;0%'
    fc[f"I{TR}"].font, fc[f"I{TR}"].border = BOLD, TOP
    fc.conditional_formatting.add(f"H{W1}:H{TR}", CellIsRule(
        operator="greaterThan", formula=["0"],
        font=Font(color="C00000", bold=True)))
    fc.conditional_formatting.add(f"H{W1}:H{TR}", CellIsRule(
        operator="lessThan", formula=["0"],
        font=Font(color="00703C", bold=True)))

    o = TR + 2  # 13
    fc[f"A{o}"] = "Overrun after the reserve"
    fc[f"A{o}"].font = SEC_FONT
    lines = [
        ("Expected final cost of work packages (EAC)", f"=F{TR}", EUR),
        ("minus current work-package budget", f"=-G{TR}", EUR),
        ("minus remaining reserve", f"=-{RESERVE_LEFT}", EUR),
        ("Overrun after using the reserve", f"=SUM(D{o + 1}:D{o + 3})",
         EUR_SIGNED),
        ("as % of the current total budget", f"=D{o + 4}/{CUR_TOTAL}", PCT1),
        ("Committed costs included in ETC", f"=E{TR}", EUR),
        ("Check: equals the Commitments sheet total",
         f'=IF(E{TR}={COMMIT_TOTAL},"OK","CHECK")', None),
    ]
    for k, (lab, f, fmt) in enumerate(lines, start=o + 1):
        cell(fc, f"B{k}", lab, bold=(k == o + 4))
        fc.merge_cells(f"B{k}:C{k}")
        cell(fc, f"D{k}", f, fmt, bold=(k == o + 4), align="right")
    ok_format(fc, f"D{o + 7}", f"D{o + 7}")
    OVERRUN, OVERRUN_PCT = f"Forecast!$D${o + 4}", f"Forecast!$D${o + 5}"
    name("OVERRUN", "Forecast", f"$D${o + 4}")
    name("OVERRUN_PCT", "Forecast", f"$D${o + 5}")
    name("COMMITTED", "Forecast", f"$D${o + 6}")
    COMMIT_CHECK = f"Forecast!$D${o + 7}"

    g = o + 9  # forecast range
    fc[f"A{g}"] = "Forecast range: how sensitive is the EAC?"
    fc[f"A{g}"].font = SEC_FONT
    header(fc, g + 1, ["", "Method", "", "EAC", "Overrun after reserve"])
    fc.merge_cells(f"B{g + 1}:C{g + 1}")
    methods = [
        ("EAC_OPT", "Optimistic: remaining work costs exactly as planned",
         f"={AC_T}+({BAC_T}-{EV_T})"),
        ("EAC_BU", "Bottom-up: team estimates to complete (used in report)",
         f"=F{TR}"),
        ("EAC_CPI", "Cost efficiency so far continues (BAC / CPI)",
         f"={BAC_T}/{CPI}"),
        ("EAC_CPISPI", "Cost and schedule pressure continue "
         "(AC + remaining / (CPI x SPI))",
         f"={AC_T}+({BAC_T}-{EV_T})/({CPI}*{SPI})"),
    ]
    for k, (nm, lab, f) in enumerate(methods, start=g + 2):
        cell(fc, f"B{k}", lab, bold=(nm == "EAC_BU"), align="left")
        fc.merge_cells(f"B{k}:C{k}")
        cell(fc, f"D{k}", f, EUR, bold=(nm == "EAC_BU"))
        cell(fc, f"E{k}", f"=D{k}-{BAC_T}-{RESERVE_LEFT}", EUR_SIGNED)
        fc.row_dimensions[k].height = 28
        if nm != "EAC_BU":
            name(nm, "Forecast", f"$D${k}")
    tp = g + 7
    tc = [
        ("TCPI_BAC", "Efficiency needed on remaining work to finish within "
         "budget (TCPI)", f"=({BAC_T}-{EV_T})/({BAC_T}-{AC_T})"),
        ("TCPI_EAC", "Efficiency needed to achieve the bottom-up EAC",
         f"=({BAC_T}-{EV_T})/(F{TR}-{AC_T})"),
        (None, "Efficiency achieved so far (CPI)", f"={CPI}"),
    ]
    for k, (nm, lab, f) in enumerate(tc, start=tp):
        cell(fc, f"B{k}", lab, align="left")
        fc.merge_cells(f"B{k}:C{k}")
        cell(fc, f"D{k}", f, RATIO, bold=True)
        fc.row_dimensions[k].height = 28
        if nm:
            name(nm, "Forecast", f"$D${k}")
    fc[f"B{tp + 3}"] = (
        f'=IF(D{tp + 1}>D{tp + 2},"The bottom-up EAC needs an efficiency of "'
        f'&TEXT(D{tp + 1},"0.00")&" on the remaining work, better than the "'
        f'&TEXT(D{tp + 2},"0.00")&" achieved so far, so treat it as the "'
        f'&"optimistic end of the range.","The bottom-up EAC is consistent "'
        f'&"with the efficiency achieved so far.")')
    fc[f"B{tp + 3}"].font = SUB
    fc[f"B{tp + 4}"] = (
        f'=IF(D{tp}>1.1,"Finishing within budget would need a TCPI of "'
        f'&TEXT(D{tp},"0.00")&", which is not realistic: a budget increase "'
        f'&"or a change order is needed.","")')
    fc[f"B{tp + 4}"].font = SUB
    widths(fc, {"A": 7, "B": 32, "C": 22, "D": 14, "E": 14, "F": 14,
                "G": 14, "H": 14, "I": 11, "J": 12, "K": 52, "L": 10})
    for n, ref in (("EAC", f"$F${TR}"), ("VAC", f"$H${TR}")):
        name(n, "Forecast", ref)
    name("FC_TABLE", "Forecast", f"$A${W1}:$K${WN}")
    EAC_T, VAC_T = f"Forecast!$F${TR}", f"Forecast!$H${TR}"
    ETC_CHECK_RNG = f"Forecast!$J${W1}:$J${WN}"
    RANK = f"Forecast!$L${W1}:$L${WN}"

    # ============================================================ Schedule
    sc = wb.create_sheet("Schedule", 8)
    title(sc, "Schedule forecast (earned schedule)",
          "Finds the point in the plan where today's earned value should "
          "have been reached, then measures how far behind that is.")
    header(sc, 4, ["Month", "Cumulative PV", "Cumulative AC",
                   "EV at status", "Month no."])
    nm_ = len(months)
    for k in range(nm_):
        r = 5 + k
        L = MCOLS[k]
        cell(sc, f"A{r}", f"=PlannedValue!{L}{h1}", MONTH_FMT)
        cell(sc, f"B{r}", f"=PlannedValue!{L}{cum}", EUR)
        cell(sc, f"C{r}", f"=ActualCost!{L}{TR + 1}", EUR)
        cell(sc, f"D{r}", f"=IF(AND(A{r}<={STATUS},EOMONTH(A{r},0)>="
             f"{STATUS}),{EV_T},NA())", EUR)
        cell(sc, f"E{r}", k + 1, align="center")
    lr = 4 + nm_
    B = f"B5:B{lr}"
    items = [
        ("Earned value at status date (EV)", f"={EV_T}", EUR, None),
        ("Months elapsed (actual time, AT)",
         f'=COUNTIF(A5:A{lr},"<="&{STATUS})', '0', "AT"),
        ("Full months of plan already earned",
         f'=COUNTIF({B},"<="&H5)', '0', None),
        ("Earned schedule (ES, months)",
         f"=IF(H7=0,H5/B5,IF(H7>={nm_},{nm_},H7+(H5-INDEX({B},H7))/"
         f"(INDEX({B},H7+1)-INDEX({B},H7))))", '0.00', "ES"),
        ("Schedule lag (months)", "=H6-H8", '0.00', None),
        ("Schedule lag (days)", f"=ROUND(H9*DAY(EOMONTH({STATUS},0)),0)",
         '0', "LAG_DAYS"),
        ("Planned launch", f"={LAUNCH}", DATE_FMT, None),
        ("Forecast launch (never before the day after the status date)",
         f"=MAX(H11+H10,{STATUS}+1)", DATE_FMT, "FORECAST_LAUNCH"),
        ("SPI (cost-based) = EV / PV", f"={SPI}", RATIO, None),
        ("SPI(t) (time-based) = ES / AT", "=H8/H6", RATIO, "SPI_T"),
        ("SPI of testing (WP5)",
         f"=INDEX(EarnedValue!$K${W1}:$K${WN},5)", RATIO, "SPI_WP5"),
        ("Forecast project end (forecast launch + aftercare weeks)",
         f"=H12+7*{AFTERCARE}", DATE_FMT, "FORECAST_END"),
    ]
    for k, (lab, f, fmt, nm) in enumerate(items, start=5):
        cell(sc, f"G{k}", lab, bold=(k == 12))
        cell(sc, f"H{k}", f, fmt, bold=(k == 12),
             fill=GREEN_LIGHT if k == 12 else None, align="right")
        if nm:
            name(nm, "Schedule", f"$H${k}")
    sc["G17"] = (f'="Lag of "&TEXT(H9,"0.00")&" months x "&DAY(EOMONTH('
                 f'{STATUS},0))&" days = "&H10&" days, so the "&TEXT(H11,'
                 f'"d mmm")&" launch moves to about "&TEXT(H12,"d mmm yyyy")'
                 f'&"."')
    sc["G18"] = ('="Assumption: the current lag does not grow. Testing '
                 '(WP5) is at SPI "&TEXT(H15,"0.00")&", so the launch date '
                 'is at risk if testing does not speed up."')
    for r in (17, 18):
        sc[f"G{r}"].font = SUB
    # plan vs actual / forecast per work package (inputs)
    ws0 = 21
    sc[f"A{ws0 - 1}"] = (f'="Plan vs actual / forecast per work package '
                         f'(as of "&TEXT({STATUS},"d mmm yyyy")&")"')
    sc[f"A{ws0 - 1}"].font = SEC_FONT
    header(sc, ws0, ["WP", "Work package", "Planned start", "Planned finish",
                     "Status at status date", "", "Forecast finish"])
    for k, (wp, (ps, pf, st, ff)) in enumerate(
            zip(s["wps"], s["wp_schedule"]), start=ws0 + 1):
        cell(sc, f"A{k}", wp[0])
        cell(sc, f"B{k}", wp[1], align="left")
        cell(sc, f"C{k}", ps, DATE_FMT, inp=True)
        cell(sc, f"D{k}", pf, DATE_FMT, inp=True)
        cell(sc, f"E{k}", st, inp=True, align="left")
        sc.merge_cells(f"E{k}:F{k}")
        cell(sc, f"G{k}", ff, inp=True, align="left")
        sc.row_dimensions[k].height = 30
    k += 1
    cell(sc, f"A{k}", "Launch")
    cell(sc, f"B{k}", "Go-live", align="left")
    cell(sc, f"D{k}", f"={LAUNCH}", DATE_FMT)
    cell(sc, f"E{k}", f'=TEXT(H10,"0")&" days late (earned schedule)"',
         align="left")
    sc.merge_cells(f"E{k}:F{k}")
    cell(sc, f"G{k}", "=H12", DATE_FMT, align="left")
    k += 1
    cell(sc, f"A{k}", "End")
    cell(sc, f"B{k}", "Project end", align="left")
    cell(sc, f"C{k}", "=Settings!$B$7", DATE_FMT)
    cell(sc, f"D{k}", f"={PLANNED_END}", DATE_FMT)
    cell(sc, f"E{k}", f'=TEXT(H16-{PLANNED_END},"0")&" days after the '
         f'planned end"', align="left")
    sc.merge_cells(f"E{k}:F{k}")
    cell(sc, f"G{k}", "=H16", DATE_FMT, align="left")
    name("WP_SCHEDULE", "Schedule", f"$A${ws0 + 1}:$G${k}")
    widths(sc, {"A": 12, "B": 26, "C": 14, "D": 14, "E": 22, "F": 18,
                "G": 42, "H": 14})
    name("SCHED_TABLE", "Schedule", f"$A$5:$D${lr}")
    FORECAST_LAUNCH, LAG_DAYS = "Schedule!$H$12", "Schedule!$H$10"
    FORECAST_END = "Schedule!$H$16"

    # ============================================================== Checks
    ck = wb.create_sheet("Checks", 9)
    title(ck, "Check figures and model integrity",
          "Part A reconciles the model to the figures in the project brief "
          "(base scenario). Part B checks the model is internally consistent "
          "for any inputs.")
    ck["A3"] = "A. Brief check figures"
    ck["A3"].font = SEC_FONT
    header(ck, 4, ["Measure", "Formula", "Model result", "Expected",
                   "Result"])
    checks = [
        ("CPI (cost efficiency)", "EV / AC", f"=ROUND({CPI},2)", 0.86, RATIO),
        ("SPI (schedule progress)", "EV / PV", f"=ROUND({SPI},2)", 0.85,
         RATIO),
        ("Earned value (EV)", "% complete x budget", f"={EV_T}", 40350, EUR),
        ("Planned value to date (PV)", "revised baseline to status date",
         f"={PV_T}", 47500, EUR),
        ("Actual cost to date (AC)", "actuals to status date", f"={AC_T}",
         46800, EUR),
        ("Cost so far vs planned", "AC - PV", f"={AC_T}-{PV_T}", -700,
         EUR_SIGNED),
        ("Expected final cost (EAC)", "AC + ETC", f"={EAC_T}", 65800, EUR),
        ("Variance at completion (VAC)", "EAC - BAC", f"={VAC_T}", 8300,
         EUR_SIGNED),
        ("Reserve remaining", "4,000 - CR-01", f"={RESERVE_LEFT}", 2500, EUR),
        ("Overrun after the reserve", "EAC - BAC - reserve", f"={OVERRUN}",
         5800, EUR_SIGNED),
        ("Overrun as % of total budget", "overrun / 60,000",
         f"=ROUND({OVERRUN_PCT},3)", 0.097, PCT1),
        ("Forecast launch date", "planned launch + schedule lag",
         f"={FORECAST_LAUNCH}", date(2024, 8, 29), DATE_FMT),
        ("Forecast project end", "forecast launch + 6 weeks aftercare",
         f"={FORECAST_END}", date(2024, 10, 10), DATE_FMT),
    ]
    for k, (m, f, model, exp, fmt) in enumerate(checks, start=5):
        cell(ck, f"A{k}", m)
        cell(ck, f"B{k}", f)
        cell(ck, f"C{k}", model, fmt, bold=True)
        cell(ck, f"D{k}", exp, fmt, inp=True)
        cell(ck, f"E{k}", f'=IF(ABS(C{k}-D{k})<0.0001,"PASS","FAIL")',
             align="center", bold=True)
    la = 4 + len(checks)
    cell(ck, f"A{la + 1}", "Brief figures reproduced", bold=True)
    cell(ck, f"E{la + 1}", f'=IF(COUNTIF(E5:E{la},"PASS")={len(checks)},'
         f'"ALL PASS","REVIEW")', align="center", bold=True)
    bi = la + 3
    ck[f"A{bi}"] = "B. Model integrity"
    ck[f"A{bi}"].font = SEC_FONT
    header(ck, bi + 1, ["Check", "Why it matters", "", "", "Result"])
    integ = [
        ("Revised PV per WP equals current budget",
         "Approved CRs must be phased into the baseline",
         f'COUNTIF({PV_CHECK_RNG},"OK")={nwp}'),
        ("Budget reconciles to contract value",
         "Original and current budgets add up to the contract",
         f'AND({BUDGET_CHECKS[1]}="OK",{BUDGET_CHECKS[2]}="OK")'),
        ("Reserve not overdrawn", "Approved CRs cannot use more reserve "
         "than exists", f'{BUDGET_CHECKS[0]}="OK"'),
        ("Approved CRs have a phasing month",
         "Otherwise PV ignores the change",
         'COUNTIF(ChangeLog!$J$5:$J$30,"Add phasing month")=0'),
        ("% complete between 0% and 100%", "EV cannot exceed budget",
         f'COUNTIFS({PCT_RNG},">=0",{PCT_RNG},"<=1")={nwp}'),
        ("ETC covers committed costs per WP",
         "Ordered costs must be in the forecast",
         f'COUNTIF({ETC_CHECK_RNG},"OK")={nwp}'),
        ("Commitments reconcile", "Forecast uses every committed cost",
         f'{COMMIT_CHECK}="OK"'),
        ("No actuals booked after the status date",
         "Future-dated costs would be ignored", AC_FUTURE_OK),
    ]
    for k, (m, why, cond) in enumerate(integ, start=bi + 2):
        cell(ck, f"A{k}", m)
        cell(ck, f"B{k}", why)
        ck.merge_cells(f"B{k}:D{k}")
        cell(ck, f"E{k}", f'=IF({cond},"PASS","FAIL")', align="center",
             bold=True)
    lb = bi + 1 + len(integ)
    cell(ck, f"A{lb + 1}", "Model integrity", bold=True)
    cell(ck, f"E{lb + 1}", f'=IF(COUNTIF(E{bi + 2}:E{lb},"PASS")='
         f'{len(integ)},"ALL PASS","REVIEW")', align="center", bold=True)
    ok_format(ck, f"E5:E{lb + 1}", "E5")
    widths(ck, {"A": 40, "B": 34, "C": 15, "D": 15, "E": 12})
    name("BRIEF_CHECKS", "Checks", f"$E${la + 1}")
    name("INTEGRITY", "Checks", f"$E${lb + 1}")
    name("CHECKS_TABLE", "Checks", f"$A$5:$E${la}")

    # ======================================================= RisksDecisions
    rk = wb.create_sheet("RisksDecisions", 10)
    title(rk, "Risks and decisions",
          "Content for the monthly report. Impact and likelihood are the "
          "project team's assessment; due dates are proposals.")
    header(rk, 4, ["#", "Top risk", "Impact", "Likelihood",
                   "Owner / response"])
    risks = [
        ("More integration bugs are found during testing (WP5)", "High",
         "Medium", "Dev lead: fix JTL-Wawi defects first; daily triage "
         "during WP5"),
        ("Data quality problems (duplicate product codes, broken umlauts) "
         "delay the launch", "High", "Medium",
         "Migration lead: UTF-8 export, de-duplication script, test "
         "re-import before go-live"),
        ("BergWerk wants Klarna (CR-02) live before launch", "Medium",
         "High", "PM: agree scope and date with BergWerk this week"),
    ]
    for k, (r, i, lk, o_) in enumerate(risks, start=5):
        cell(rk, f"A{k}", k - 4, align="center")
        cell(rk, f"B{k}", r, inp=True, align="left")
        cell(rk, f"C{k}", i, inp=True, align="center")
        cell(rk, f"D{k}", lk, inp=True, align="center")
        cell(rk, f"E{k}", o_, inp=True, align="left")
        rk.row_dimensions[k].height = 32
    cell(rk, "A8", "W", align="center")
    cell(rk, "B8", f'="Watch: only EUR "&TEXT({RESERVE_LEFT},"#,##0")&" of '
         f'reserve is left; once it goes to WP3 there is no contingency for '
         f'the last two months"', align="left")
    cell(rk, "C8", "High", inp=True, align="center")
    cell(rk, "D8", "High", inp=True, align="center")
    cell(rk, "E8", "PM: any new issue needs a change order or a budget "
         "increase", inp=True, align="left")
    rk.row_dimensions[8].height = 32
    header(rk, 10, ["#", "Decision needed from the project manager",
                    "Cost / date effect", "Proposed due", "Recommendation"])
    decisions = [
        ("Approve CR-02 (Klarna) now, or move it to after launch",
         f'=SUMIFS({CL_AMT},{CL_ID},"CR-02")', EUR, date(2024, 8, 2),
         "Move to after launch: keeps WP5 testing focused; BergWerk pays "
         "EUR 2,500 either way"),
        ("Spend the remaining reserve on WP3", f"={RESERVE_LEFT}", EUR,
         date(2024, 8, 2), "Approve: WP3 is the largest overrun, but this "
         "leaves no contingency"),
        ("Ask BergWerk to pay for the extra integration work through a "
         "change order", "=INDEX(Forecast!$H$5:$H$10,3)", EUR,
         date(2024, 8, 9), "First check the contract: if JTL-Wawi API "
         "defects are outside the agreed assumptions, raise a change order "
         "with the defect log as evidence; otherwise absorb the cost"),
        ("Agree the new launch date with BergWerk", f"={FORECAST_LAUNCH}",
         DATE_FMT, date(2024, 8, 2), "Confirm the forecast launch date in "
         "writing and re-plan WP6 support"),
    ]
    for k, (d_, eff, fmt, due, rec) in enumerate(decisions, start=11):
        cell(rk, f"A{k}", k - 10, align="center")
        cell(rk, f"B{k}", d_, inp=True, align="left")
        cell(rk, f"C{k}", eff, fmt, align="center")
        cell(rk, f"D{k}", due, DATE_FMT, inp=True, align="center")
        cell(rk, f"E{k}", rec, inp=True, align="left")
        rk.row_dimensions[k].height = 45
    widths(rk, {"A": 4, "B": 58, "C": 14, "D": 13, "E": 64})
    name("RISKS", "RisksDecisions", "$B$5:$E$8")
    name("DECISIONS", "RisksDecisions", "$B$11:$E$14")

    # =========================================================== Dashboard
    db = wb.create_sheet("Dashboard", 0)
    title(db, "BergWerk Outdoor: Shopify migration, cost report",
          "Monthly status at the status date. All figures in EUR and linked "
          "to the detail sheets.")
    cell(db, "A3", "Status date", border=False)
    cell(db, "B3", f"={STATUS}", DATE_FMT, bold=True, border=False)
    cell(db, "C3", '="Model integrity: "&Checks!$E$' + str(lb + 1) +
         '&"   |   Brief figures: "&Checks!$E$' + str(la + 1), border=False)
    header(db, 5, ["Status", "Rating", "Rule"])
    cell(db, "A6", "Cost")
    cell(db, "B6", f'=IF({OVERRUN}>0,"RED",IF({VAC_T}>0,"AMBER","GREEN"))',
         bold=True, align="center")
    cell(db, "C6", "RED = forecast overrun larger than the remaining "
         "reserve", align="left")
    cell(db, "A7", "Schedule")
    cell(db, "B7", f'=IF({SPI}>={SPI_G},"GREEN",IF({SPI}>={SPI_A},'
         f'"AMBER","RED"))', bold=True, align="center")
    cell(db, "C7", '="GREEN if SPI at least "&TEXT(' + SPI_G +
         ',"0.00")&", AMBER if at least "&TEXT(' + SPI_A + ',"0.00")',
         align="left")
    name("COST_RAG", "Dashboard", "$B$6")
    name("SCHED_RAG", "Dashboard", "$B$7")
    for color, fill, font in (("RED", "C00000", "FFFFFF"),
                              ("AMBER", "F4B183", "000000"),
                              ("GREEN", "70AD47", "FFFFFF")):
        db.conditional_formatting.add("B6:B7", CellIsRule(
            operator="equal", formula=[f'"{color}"'],
            fill=PatternFill("solid", fgColor=fill),
            font=Font(color=font, bold=True)))

    header(db, 9, ["Key figure", "Value", "Meaning"])
    kpis = [
        ("Budget (work packages, after approved CRs)", f"={BAC_T}", EUR,
         "Budget at completion (BAC)"),
        ("Planned value to date (PV)", f"={PV_T}", EUR,
         "Work that should have been done by the status date"),
        ("Earned value (EV)", f"={EV_T}", EUR,
         "Budget value of the work actually done"),
        ("Actual cost to date (AC)", f"={AC_T}", EUR,
         "What the work has cost"),
        ("Cost so far vs planned (AC - PV)", f"={AC_T}-{PV_T}", EUR_SIGNED,
         "Looks fine on the surface, which is why earned value is needed"),
        ("CPI = EV / AC", f"={CPI}", RATIO,
         f'=TEXT({CPI}*100,"0")&" cents of planned work per euro spent"'),
        ("SPI = EV / PV", f"={SPI}", RATIO,
         f'=TEXT((1-{SPI})*100,"0")&"% behind plan"'),
        ("Expected final cost (EAC, bottom-up)", f"={EAC_T}", EUR,
         '="Range "&TEXT(Forecast!$D$' + str(g + 2) + ',"#,##0")&" to "&'
         'TEXT(Forecast!$D$' + str(g + 5) + ',"#,##0")&" depending on '
         'method"'),
        ("Variance at completion (EAC - BAC)", f"={VAC_T}", EUR_SIGNED,
         "Forecast overrun against the work-package budget"),
        ("Reserve remaining", f"={RESERVE_LEFT}", EUR,
         "After approved CRs paid from the reserve"),
        ("Overrun after using the reserve", f"={OVERRUN}", EUR_SIGNED,
         f'=TEXT({OVERRUN_PCT},"0.0%")&" of the EUR "&TEXT({CUR_TOTAL},'
         f'"#,##0")&" total budget"'),
        ("TCPI to finish within budget", f"=Forecast!$D${tp}", RATIO,
         "Efficiency the remaining work would need; above 1.10 is not "
         "realistic"),
        ("Planned launch", f"={LAUNCH}", DATE_FMT, "Original go-live date"),
        ("Forecast launch", f"={FORECAST_LAUNCH}", DATE_FMT,
         f'="About "&{LAG_DAYS}&" days late (earned schedule)"'),
        ("Planned project end", f"={PLANNED_END}", DATE_FMT,
         "End of aftercare in the original plan"),
        ("Forecast project end", f"={FORECAST_END}", DATE_FMT,
         f'="Forecast launch + "&{AFTERCARE}&" weeks aftercare"'),
    ]
    for k, (lab, f, fmt, meaning) in enumerate(kpis, start=10):
        cell(db, f"A{k}", lab)
        cell(db, f"B{k}", f, fmt, bold=True, align="right")
        cell(db, f"C{k}", meaning, align="left")
    dh = 10 + len(kpis) + 1
    header(db, dh, ["Main differences from budget", "EAC - BAC", "Reason"])
    picks = [f"LARGE({RANK},{k})" for k in (1, 2, 3)] + [f"SMALL({RANK},1)"]
    for k, pick in enumerate(picks, start=dh + 1):
        idx = f"MATCH({pick},{RANK},0)"
        cell(db, f"A{k}", f'=INDEX(Forecast!$A$5:$A$10,{idx})&" "&'
             f'INDEX(Forecast!$B$5:$B$10,{idx})&" ("&TEXT(INDEX('
             f'Forecast!$I$5:$I$10,{idx}),"+0%;-0%;0%")&")"')
        cell(db, f"B{k}", f"=INDEX(Forecast!$H$5:$H$10,{idx})", EUR_SIGNED,
             bold=True, align="right")
        cell(db, f"C{k}", f"=INDEX(Forecast!$K$5:$K$10,{idx})", align="left")
        db.row_dimensions[k].height = 28
    db[f"A{dh + 6}"] = ("Ranked automatically: the three largest overruns, "
                        "then the largest underspend.")
    db[f"A{dh + 6}"].font = SUB
    name("DIFF_TABLE", "Dashboard", f"$A${dh + 1}:$C${dh + 4}")
    widths(db, {"A": 42, "B": 16, "C": 62})

    line = LineChart()
    line.title = "Cumulative PV, AC and EV (EUR)"
    line.height, line.width = 8, 16
    line.y_axis.title = "EUR"
    line.y_axis.numFmt = '#,##0'
    line.add_data(Reference(sc, min_col=2, max_col=4, min_row=4, max_row=lr),
                  titles_from_data=True)
    line.set_categories(Reference(sc, min_col=1, min_row=5, max_row=lr))
    for sr, colr in zip(line.series, ("7F7F7F", "C00000", "1F4E3D")):
        sr.graphicalProperties.line.solidFill = colr
        sr.graphicalProperties.line.width = 28000
        sr.smooth = False
    line.series[0].graphicalProperties.line.dashStyle = "dash"
    line.series[2].marker = Marker(symbol="diamond", size=10)
    line.series[2].marker.graphicalProperties.solidFill = "1F4E3D"
    line.x_axis.number_format = 'mmm'
    line.x_axis.delete = line.y_axis.delete = False
    db.add_chart(line, "E5")

    bar = BarChart()
    bar.type = "col"
    bar.title = "Budget vs expected final cost by work package"
    bar.height, bar.width = 8, 16
    bar.add_data(Reference(fc, min_col=6, max_col=7, min_row=4, max_row=WN),
                 titles_from_data=True)
    bar.set_categories(Reference(fc, min_col=1, min_row=5, max_row=WN))
    bar.series[0].graphicalProperties.solidFill = "C00000"
    bar.series[1].graphicalProperties.solidFill = "A9C5B4"
    bar.y_axis.numFmt = '#,##0'
    bar.x_axis.delete = bar.y_axis.delete = False
    db.add_chart(bar, "E22")

    for n, ref in names.items():
        wb.defined_names[n] = DefinedName(n, attr_text=ref)
    for w in wb.worksheets:
        w.sheet_properties.pageSetUpPr.fitToPage = True
        w.page_setup.orientation = "landscape"
        w.page_setup.fitToWidth = 1
        w.page_setup.fitToHeight = 0
    db.sheet_properties.tabColor = GREEN_DARK
    ck.sheet_properties.tabColor = "70AD47"
    wb.calculation.fullCalcOnLoad = True  # Excel recalculates on open
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out


if __name__ == "__main__":
    print(f"Saved {build()}")

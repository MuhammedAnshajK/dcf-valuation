"""Builds dcf_valuation.xlsx: a DCF (perpetuity growth and exit multiple) and
comparable-companies valuation of the illustrative company from the
three-statement-model project. Run: python build_model.py"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList

wb = Workbook(); F = "Arial"
def fnt(color="000000", bold=False, italic=False, size=10): return Font(name=F, color=color, bold=bold, italic=italic, size=size)
BLUE, BLACK, GREEN, GREY = "0000FF", "000000", "008000", "595959"
HFILL = PatternFill("solid", fgColor="1F4E78"); YEL = PatternFill("solid", fgColor="FFFF00")
LIGHT = PatternFill("solid", fgColor="DDEBF7")
TOP = Border(top=Side(style="thin")); DBL = Border(top=Side(style="thin"), bottom=Side(style="double"))
NUM = '#,##0;(#,##0);"-"'; PCT = '0.0%;(0.0%);"-"'; MULT = '0.0"x"'; PS = '$#,##0.00;($#,##0.00);"-"'; DEC = '0.00'

def cell(ws, ref, v, fmt=None, color=None, bold=False, fill=None, border=None, italic=False):
    c = ws[ref]; c.value = v
    if color is None:
        color = (GREEN if "!" in v else BLACK) if isinstance(v, str) and v.startswith("=") else (BLUE if isinstance(v, (int, float)) else BLACK)
    c.font = fnt(color, bold, italic)
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    if border: c.border = border
    return c

def title(ws, t, sub="Illustrative Manufacturing Co. ($ thousands unless stated)"):
    ws["B1"] = t; ws["B1"].font = fnt(bold=True, size=14)
    ws["B2"] = sub; ws["B2"].font = fnt(GREY, italic=True)
    ws.column_dimensions["A"].width = 2; ws.column_dimensions["B"].width = 40
    ws.sheet_view.showGridLines = False

def header(ws, r, labels, start_col=3):
    ws[f"B{r}"].fill = HFILL
    for i, lab in enumerate(labels):
        col = chr(64 + start_col + i); c = ws[f"{col}{r}"]; c.value = lab
        c.font = fnt("FFFFFF", True); c.fill = HFILL; c.alignment = Alignment(horizontal="right")
        ws.column_dimensions[col].width = 12

YRS = ["2025A", "2026E", "2027E", "2028E", "2029E", "2030E"]; YC = list("CDEFGH"); PC = YC[1:]
def prv(c): return chr(ord(c) - 1)

# ---------------- Cover ----------------
cv = wb.active; cv.title = "Cover"; cv.sheet_view.showGridLines = False
cv.column_dimensions["A"].width = 2; cv.column_dimensions["B"].width = 22; cv.column_dimensions["C"].width = 95
cv["B2"] = "DCF & Comparable Companies Valuation"; cv["B2"].font = fnt(bold=True, size=18)
cv["B3"] = "Illustrative Manufacturing Co. | valuation date 31 Dec 2025 | $ thousands unless stated"; cv["B3"].font = fnt(GREY, italic=True)
info = [
 ("Purpose", "Values the company three ways (DCF with perpetuity growth, DCF with an exit multiple, and trading comparables) and compares them on one summary page."),
 ("Data", "Forecast from the base case of the three-statement-model project. Market inputs and peer multiples are illustrative, not real market data."),
 ("How to use", "Change the blue or yellow cells on the WACC, DCF and Comps sheets. Everything else is a formula."),
 ("Sheets", "Forecast > WACC > DCF > Sensitivity > Comps > Summary"),
 ("Colour key", "Blue = hardcoded input | Black = formula | Green = link to another sheet | Yellow fill = key assumption"),
 ("Author", "Muhammed Anshaj K"),
]
for i, (k, v) in enumerate(info):
    cv[f"B{5+i}"] = k; cv[f"B{5+i}"].font = fnt(bold=True); cv[f"B{5+i}"].alignment = Alignment(vertical="top")
    cv[f"C{5+i}"] = v; cv[f"C{5+i}"].font = fnt(); cv[f"C{5+i}"].alignment = Alignment(wrap_text=True, vertical="top")

# ---------------- Forecast (from three-statement model) ----------------
fc = wb.create_sheet("Forecast"); title(fc, "Forecast inputs"); header(fc, 4, YRS)
data = {
 5: ("Revenue", [12000, 13200, 14388, 15539.0, 16626.8, 17624.4]),
 6: ("EBITDA", [2340, 2640, 2949.5, 3340.9, 3574.8, 3877.4]),
 7: ("Depreciation & amortisation", [600, 660, 719.4, 777.0, 831.3, 881.2]),
 9: ("Capital expenditure", [None, 858, 863.3, 932.3, 914.5, 969.3]),
 10: ("Net working capital (AR + inventory - AP)", [1940, 2133.7, 2322.8, 2505.4, 2680.8, 2841.6]),
 11: ("Net income", [1155, None, None, None, None, None]),
}
for r, (lab, vals) in data.items():
    cell(fc, f"B{r}", lab)
    for c, v in zip(YC, vals):
        if v is not None: cell(fc, f"{c}{r}", v, NUM)
cell(fc, "B8", "EBIT", bold=True)
for c in YC: cell(fc, f"{c}8", f"={c}6-{c}7", NUM, bold=True, border=TOP)
cell(fc, "B13", "Balance sheet at 31 Dec 2025", bold=True)
cell(fc, "B14", "Cash"); cell(fc, "C14", 908, NUM)
cell(fc, "B15", "Debt"); cell(fc, "C15", 3000, NUM)
cell(fc, "B16", "Shares outstanding (thousands)"); cell(fc, "C16", 1000, NUM, fill=YEL)
cell(fc, "B18", "Source: base case of the three-statement-model project (github.com/MuhammedAnshajK/three-statement-model), rounded to 0.1. Shares outstanding are illustrative.", italic=True, color=GREY)
fc["C16"].comment = Comment("Illustrative share count: 1,000 thousand shares (1.0m).", "Model")

# ---------------- WACC ----------------
w = wb.create_sheet("WACC"); title(w, "Discount rate (WACC)"); w.column_dimensions["C"].width = 12; w.column_dimensions["D"].width = 60
rows = [
 (5, "Risk-free rate", 0.045, PCT, "Illustrative long-term government bond yield"),
 (6, "Equity risk premium", 0.055, PCT, "Illustrative market premium"),
 (7, "Beta (levered)", 1.10, DEC, "Illustrative, in line with industrial peers"),
 (8, "Cost of equity (CAPM)", "=C5+C7*C6", PCT, "Risk-free + beta x equity risk premium"),
 (10, "Pre-tax cost of debt", 0.07, PCT, "Matches the 7% interest rate in the three-statement model"),
 (11, "Tax rate", 0.25, PCT, "Matches the forecast tax rate"),
 (12, "After-tax cost of debt", "=C10*(1-C11)", PCT, ""),
 (14, "Target debt / (debt + equity)", 0.20, PCT, "Illustrative target capital structure"),
 (15, "Target equity / (debt + equity)", "=1-C14", PCT, ""),
 (17, "WACC", "=C15*C8+C14*C12", PCT, "Weighted average cost of capital"),
]
for r, lab, v, fmt, note in rows:
    cell(w, f"B{r}", lab, bold=(r in (8, 17)))
    cell(w, f"C{r}", v, fmt, bold=(r in (8, 17)), fill=(YEL if r in (5, 6, 7, 14) else (LIGHT if r == 17 else None)),
         border=(DBL if r == 17 else None))
    cell(w, f"D{r}", note, italic=True, color=GREY)

# ---------------- DCF ----------------
d = wb.create_sheet("DCF"); title(d, "Discounted cash flow"); header(d, 4, YRS)
cell(d, "B5", "EBIT")
for c in YC: cell(d, f"{c}5", f"=Forecast!{c}8", NUM)
cell(d, "B6", "Less: tax on EBIT")
for c in PC: cell(d, f"{c}6", f"=-{c}5*WACC!$C$11", NUM)
cell(d, "B7", "NOPAT", bold=True)
for c in PC: cell(d, f"{c}7", f"={c}5+{c}6", NUM, bold=True, border=TOP)
cell(d, "B8", "Add: depreciation & amortisation")
for c in PC: cell(d, f"{c}8", f"=Forecast!{c}7", NUM)
cell(d, "B9", "Less: capital expenditure")
for c in PC: cell(d, f"{c}9", f"=-Forecast!{c}9", NUM)
cell(d, "B10", "Less: increase in net working capital")
for c in PC: cell(d, f"{c}10", f"=-(Forecast!{c}10-Forecast!{prv(c)}10)", NUM)
cell(d, "B11", "Unlevered free cash flow", bold=True)
for c in PC: cell(d, f"{c}11", f"=SUM({c}7:{c}10)", NUM, bold=True, border=TOP)
cell(d, "B13", "Discount period (years)")
for i, c in enumerate(PC): cell(d, f"{c}13", i + 1, '0')
cell(d, "B14", "Discount factor")
for c in PC: cell(d, f"{c}14", f"=1/(1+$C$18)^{c}13", '0.000')
cell(d, "B15", "Present value of free cash flow")
for c in PC: cell(d, f"{c}15", f"={c}11*{c}14", NUM)

cell(d, "B17", "Assumptions", bold=True)
cell(d, "B18", "WACC"); cell(d, "C18", "=WACC!C17", PCT)
cell(d, "B19", "Terminal growth rate"); cell(d, "C19", 0.025, PCT, fill=YEL)
cell(d, "B20", "Exit EV / EBITDA multiple"); cell(d, "C20", 8.0, MULT, fill=YEL)

# Perpetuity growth method
cell(d, "B22", "Valuation", bold=True)
cell(d, "C22", "Perpetuity", bold=True); cell(d, "E22", "Exit multiple", bold=True)
for ref in ("C22", "E22"): d[ref].alignment = Alignment(horizontal="right")
rowsP = [
 (23, "Sum of PV of free cash flows", "=SUM(D15:H15)", "=SUM(D15:H15)"),
 (24, "Terminal value (2030)", "=H11*(1+C19)/(C18-C19)", "=Forecast!H6*C20"),
 (25, "PV of terminal value", "=C24*H14", "=E24*H14"),
 (26, "Enterprise value", "=C23+C25", "=E23+E25"),
 (27, "Less: debt", "=-Forecast!C15", "=-Forecast!C15"),
 (28, "Add: cash", "=Forecast!C14", "=Forecast!C14"),
 (29, "Equity value", "=SUM(C26:C28)", "=SUM(E26:E28)"),
 (30, "Value per share ($)", "=C29/Forecast!C16", "=E29/Forecast!C16"),
 (32, "Terminal value as % of EV", "=C25/C26", "=E25/E26"),
 (33, "Implied exit multiple / implied growth", "=C24/Forecast!H6", "=(E24*C18-H11)/(E24+H11)"),
]
for r, lab, p, e in rowsP:
    bold = r in (26, 29, 30)
    fmt = PS if r == 30 else (PCT if r == 32 else NUM)
    cell(d, f"B{r}", lab, bold=bold)
    cell(d, f"C{r}", p, MULT if r == 33 else fmt, bold=bold, border=(TOP if r in (26, 29) else None))
    cell(d, f"E{r}", e, PCT if r == 33 else fmt, bold=bold, border=(TOP if r in (26, 29) else None))
for r in (30,): d[f"C{r}"].fill = LIGHT; d[f"E{r}"].fill = LIGHT
cell(d, "B35", "Cash flows are discounted at year end. Terminal value uses the 2030 free cash flow grown one year.", italic=True, color=GREY)

# ---------------- Sensitivity ----------------
s = wb.create_sheet("Sensitivity"); title(s, "Sensitivity: value per share ($)", "Perpetuity growth method | WACC down the side, terminal growth across the top")
s.column_dimensions["B"].width = 18
waccs = [-0.01, -0.005, 0, 0.005, 0.01]; gs = [-0.01, -0.005, 0, 0.005, 0.01]
cell(s, "B4", "WACC \\ growth", bold=True)
for j, dg in enumerate(gs):
    col = chr(ord("C") + j); s.column_dimensions[col].width = 12
    cell(s, f"{col}4", f"=DCF!$C$19+({dg})", PCT, bold=True, fill=LIGHT)
for i, dw in enumerate(waccs):
    r = 5 + i
    cell(s, f"B{r}", f"=DCF!$C$18+({dw})", PCT, bold=True, fill=LIGHT)
    for j in range(5):
        col = chr(ord("C") + j)
        f = (f"=(SUMPRODUCT(DCF!$D$11:$H$11/(1+$B{r})^DCF!$D$13:$H$13)"
             f"+DCF!$H$11*(1+{col}$4)/($B{r}-{col}$4)/(1+$B{r})^DCF!$H$13"
             f"-Forecast!$C$15+Forecast!$C$14)/Forecast!$C$16")
        cell(s, f"{col}{r}", f, PS)
s["E7"].fill = YEL
cell(s, "B11", "Centre cell (yellow) equals the base case on the DCF sheet.", italic=True, color=GREY)
cell(s, "B12", "Check: centre less DCF value per share"); cell(s, "E12", "=ROUND(E7-DCF!C30,6)", PS)

# ---------------- Comps ----------------
cp = wb.create_sheet("Comps"); title(cp, "Trading comparables", "Illustrative peer group | multiples are made up for this example")
cp.column_dimensions["B"].width = 26
header(cp, 4, ["EV / Revenue", "EV / EBITDA", "P / E"])
peers = [("Peer A (industrial machinery)", 1.4, 8.5, 14.2), ("Peer B (components)", 1.1, 7.2, 12.5),
         ("Peer C (automation)", 1.9, 9.8, 17.0), ("Peer D (metal products)", 1.0, 6.9, 11.8),
         ("Peer E (equipment)", 1.3, 8.1, 13.9), ("Peer F (precision parts)", 2.1, 10.4, 18.5)]
for i, (n, a, b_, c_) in enumerate(peers):
    r = 5 + i; cell(cp, f"B{r}", n)
    for col, v in zip("CDE", (a, b_, c_)): cell(cp, f"{col}{r}", v, MULT)
stats = [(12, "25th percentile", "QUARTILE({c}5:{c}10,1)"), (13, "Median", "MEDIAN({c}5:{c}10)"), (14, "75th percentile", "QUARTILE({c}5:{c}10,3)")]
for r, lab, f in stats:
    cell(cp, f"B{r}", lab, bold=True)
    for col in "CDE": cell(cp, f"{col}{r}", "=" + f.format(c=col), MULT, bold=True, border=(TOP if r == 12 else None))
cell(cp, "B16", "Company metric (2025A)", bold=True)
cell(cp, "C16", "=Forecast!C5", NUM); cell(cp, "D16", "=Forecast!C6", NUM); cell(cp, "E16", "=Forecast!C11", NUM)
cell(cp, "B18", "Implied value per share ($)", bold=True)
for r, lab in ((19, "At 25th percentile"), (20, "At median"), (21, "At 75th percentile")):
    cell(cp, f"B{r}", lab); sr = r - 7
    for col in "CD":  # EV multiples: subtract net debt
        cell(cp, f"{col}{r}", f"=({col}{sr}*{col}$16-Forecast!$C$15+Forecast!$C$14)/Forecast!$C$16", PS)
    cell(cp, f"E{r}", f"=E{sr}*E$16/Forecast!$C$16", PS)

# ---------------- Summary / football field ----------------
sm = wb.create_sheet("Summary"); title(sm, "Valuation summary", "Value per share ($) | the bars in the chart run from low to high")
sm.column_dimensions["B"].width = 34
header(sm, 4, ["Low", "High", "Range", "Base"])
summ = [
 (5, "DCF: perpetuity growth (sensitivity)", "=MIN(Sensitivity!C5:G9)", "=MAX(Sensitivity!C5:G9)", "=DCF!C30"),
 (6, "DCF: exit multiple (7.0x to 9.0x)", "=(DCF!C23+(DCF!C20-1)*Forecast!H6*DCF!H14-Forecast!C15+Forecast!C14)/Forecast!C16",
     "=(DCF!C23+(DCF!C20+1)*Forecast!H6*DCF!H14-Forecast!C15+Forecast!C14)/Forecast!C16", "=DCF!E30"),
 (7, "Comps: EV / EBITDA (25th to 75th pct)", "=Comps!D19", "=Comps!D21", "=Comps!D20"),
 (8, "Comps: EV / Revenue (25th to 75th pct)", "=Comps!C19", "=Comps!C21", "=Comps!C20"),
 (9, "Comps: P / E (25th to 75th pct)", "=Comps!E19", "=Comps!E21", "=Comps!E20"),
]
for r, lab, lo, hi, base in summ:
    cell(sm, f"B{r}", lab); cell(sm, f"C{r}", lo, PS); cell(sm, f"D{r}", hi, PS)
    cell(sm, f"E{r}", f"=D{r}-C{r}", PS); cell(sm, f"F{r}", base, PS)
cell(sm, "B11", "Midpoint of all ranges", bold=True)
cell(sm, "C11", "=AVERAGE(C5:C9)", PS, bold=True); cell(sm, "D11", "=AVERAGE(D5:D9)", PS, bold=True)
cell(sm, "F11", "=(C11+D11)/2", PS, bold=True, fill=LIGHT)
cell(sm, "B13", "The DCF sits above the trading comps because the forecast assumes margins expand from 19.5% to 22.0% EBITDA. Peers are valued on today's earnings.", italic=True, color=GREY)

ch = BarChart(); ch.type = "bar"; ch.grouping = "stacked"; ch.overlap = 100
ch.title = "Football field: value per share ($)"; ch.y_axis.title = "$ per share"
ch.add_data(Reference(sm, min_col=3, min_row=4, max_row=9), titles_from_data=True)
ch.add_data(Reference(sm, min_col=5, min_row=4, max_row=9), titles_from_data=True)
ch.set_categories(Reference(sm, min_col=2, min_row=5, max_row=9))
ch.series[0].graphicalProperties.noFill = True; ch.series[0].graphicalProperties.line.noFill = True
ch.series[1].graphicalProperties.solidFill = "1F4E78"
ch.legend = None; ch.x_axis.scaling.orientation = "maxMin"; ch.y_axis.majorGridlines = None
ch.x_axis.delete = False; ch.y_axis.delete = False; ch.y_axis.number_format = '$#,##0'
ch.width = 22; ch.height = 9
sm.add_chart(ch, "B15")

for ws in wb.worksheets:
    ws.page_setup.orientation = "landscape"; ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 1
wb.save("dcf_valuation.xlsx")
print("Saved dcf_valuation.xlsx. Open it in Excel (or run a recalculation) to see values.")

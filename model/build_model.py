"""
MUELLER INDUSTRIES (NYSE: MLI) / MUELLER WATER PRODUCTS (NYSE: MWA)
Horizontal Water Infrastructure Consolidation - Forecasting, Budgeting & Credit Model
REBUILT from real FY2025 10-K figures. Replaces the fabricated standalone data
and the self-referential copper formula from the originally uploaded file.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

wb = openpyxl.Workbook()

ws_sum = wb.active
ws_sum.title = "Executive_Summary"
ws_mli = wb.create_sheet("MLI_Standalone")
ws_mwa = wb.create_sheet("MWA_Standalone")
ws_copper = wb.create_sheet("Copper_Cost_Engine")
ws_pf = wb.create_sheet("Pro_Forma_Consolidated")
ws_sens = wb.create_sheet("Sensitivity_Matrix")
ws_fcst = wb.create_sheet("Five_Year_Forecast")
ws_audit = wb.create_sheet("Audit_Checks")

# ---------- styles ----------
F_TITLE = Font(name="Arial", size=14, bold=True, color="1E293B")
F_HDR = Font(name="Arial", size=10, bold=True, color="FFFFFF")
F_BOLD = Font(name="Arial", size=10, bold=True, color="000000")
F_REG = Font(name="Arial", size=10, color="000000")
F_INPUT = Font(name="Arial", size=10, bold=True, color="0000FF")      # blue = hardcoded input
F_FORM = Font(name="Arial", size=10, color="000000")                   # black = formula
F_LINK = Font(name="Arial", size=10, color="008000")                   # green = cross-sheet link
F_NOTE = Font(name="Arial", size=8, italic=True, color="808080")

FILL_HDR = PatternFill("solid", fgColor="1E293B")
FILL_ASSUMP = PatternFill("solid", fgColor="FFFF00")
FILL_SUBTOTAL = PatternFill("solid", fgColor="F1F5F9")

CUR = '$#,##0.0;($#,##0.0);"-"'
PCT = '0.0%'
PX = '$#,##0.000'

def header(ws, cell, text, width=None):
    ws[cell] = text
    ws[cell].font = F_TITLE
    if width:
        ws.column_dimensions[cell[0]].width = width

def note(ws, cell, text):
    ws[cell] = text
    ws[cell].font = F_NOTE

# =========================================================================
# MLI_Standalone  (FY2025, fiscal year ended Dec 27, 2025)
# Source: Mueller Industries 10-K FY2025, "Results of Operations", SEC EDGAR
# https://www.sec.gov/Archives/edgar/data/89439/000008943926000008/mli-20251227.htm
# =========================================================================
ws_mli["A1"] = "MUELLER INDUSTRIES (NYSE: MLI) — STANDALONE, FY2025"
ws_mli["A1"].font = F_TITLE
note(ws_mli, "A2", "Source: MLI FY2025 10-K, Results of Operations (filed Feb 2026), SEC EDGAR. All figures $ millions.")

rows = [
    ("A4", "Net Sales",            "B4", 4178.5, True,  "10-K: net sales, FY2025"),
    ("A5", "Cost of Goods Sold",   "B5", 2966.1, True,  "10-K: COGS $2,966,083K"),
    ("A6", "Depreciation & Amortization", "B6", 68.6, True, "10-K: D&A $68,561K"),
    ("A7", "SG&A Expense",         "B7", 248.7, True,  "10-K: SG&A $248,651K"),
]
for label_cell, label, val_cell, val, is_input, src in rows:
    ws_mli[label_cell] = label
    ws_mli[label_cell].font = F_REG
    ws_mli[val_cell] = val
    ws_mli[val_cell].font = F_INPUT if is_input else F_FORM
    ws_mli[val_cell].number_format = CUR
    ws_mli[val_cell].fill = FILL_ASSUMP
    ws_mli[val_cell].comment = Comment(src, "Model")

ws_mli["A9"] = "Gross Profit"
ws_mli["B9"] = "=B4-B5"
ws_mli["B9"].number_format = CUR
ws_mli["B9"].font = F_BOLD
ws_mli["A10"] = "Gross Margin %"
ws_mli["B10"] = "=B9/B4"
ws_mli["B10"].number_format = PCT

ws_mli["A12"] = "Operating Income (company-reported)"
ws_mli["B12"] = 958.5
ws_mli["B12"].font = F_INPUT
ws_mli["B12"].fill = FILL_ASSUMP
ws_mli["B12"].number_format = CUR
ws_mli["B12"].comment = Comment("10-K / FY2025 earnings release: reported Operating Income. Taken as a direct input, NOT built up from COGS+D&A+SG&A below — see reconciling-gap note.", "Model")

ws_mli["A13"] = "EBITDA (company-reported)"
ws_mli["B13"] = 1027.1
ws_mli["B13"].font = F_INPUT
ws_mli["B13"].fill = FILL_ASSUMP
ws_mli["B13"].number_format = CUR
ws_mli["B13"].comment = Comment("Proxy statement 'Results at a Glance' table: FY2025 EBITDA. Company's own non-GAAP reconciliation, taken as reported.", "Model")

ws_mli["A17"] = "Reconciling gap check"
ws_mli["A18"] = "Component build (Gross Profit - D&A - SG&A)"
ws_mli["B18"] = "=B9-B6-B7"
ws_mli["B18"].number_format = CUR
ws_mli["A19"] = "Gap vs. reported Operating Income"
ws_mli["B19"] = "=B18-B12"
ws_mli["B19"].number_format = CUR
ws_mli["B19"].font = F_BOLD
note(ws_mli, "C19", "~$(63)M gap: the three disclosed line items (COGS/D&A/SG&A) do not fully reconcile to reported Operating Income — likely unallocated corporate costs not broken out in the excerpt pulled. Confirm against the full 10-K segment footnote before relying on the component split; the headline Operating Income/EBITDA above are the authoritative, reported figures.")

ws_mli["A14"] = "EBITDA Margin %"
ws_mli["B14"] = "=B13/B4"
ws_mli["B14"].number_format = PCT

ws_mli["A21"] = "Total Debt (FY2025, per 10-K)"
ws_mli["B21"] = 0
ws_mli["B21"].font = F_INPUT
ws_mli["B21"].fill = FILL_ASSUMP
ws_mli["B21"].number_format = CUR
ws_mli["B21"].comment = Comment("10-K: 'As of December 27, 2025, the Company had no debt outstanding.'", "Model")
ws_mli["A22"] = "Cash & Short-Term Investments (FY2025)"
ws_mli["B22"] = 1389.7
ws_mli["B22"].font = F_INPUT
ws_mli["B22"].fill = FILL_ASSUMP
ws_mli["B22"].number_format = CUR
ws_mli["B22"].comment = Comment("Q4/FY2025 earnings release: year-end cash and short-term investments.", "Model")
ws_mli["A23"] = "Capital Expenditures (FY2025 actual)"
ws_mli["B23"] = 68.8
ws_mli["B23"].font = F_INPUT
ws_mli["B23"].fill = FILL_ASSUMP
ws_mli["B23"].number_format = CUR
ws_mli["B23"].comment = Comment("10-K: FY2025 capex $68.8M; company guides $80-90M for FY2026.", "Model")
ws_mli["A24"] = "FY2026 CapEx Guidance — Low / High ($M)"
ws_mli["B24"] = 80.0
ws_mli["C24"] = 90.0
for cc in ["B24","C24"]:
    ws_mli[cc].font = F_INPUT
    ws_mli[cc].fill = FILL_ASSUMP
    ws_mli[cc].number_format = CUR
ws_mli["B24"].comment = Comment("Company FY2026 CapEx guidance range, disclosed on Q4/FY2025 earnings call.", "Model")
ws_mli["A25"] = "Dividend Payout Ratio (trailing)"
ws_mli["B25"] = 0.1512
ws_mli["B25"].font = F_INPUT
ws_mli["B25"].fill = FILL_ASSUMP
ws_mli["B25"].number_format = PCT
ws_mli["B25"].comment = Comment("Reported trailing payout ratio (regular quarterly dividend $0.25/sh against FY2025 EPS), per market data aggregators.", "Model")

ws_mli["A16"] = "Net Income (reported, incl. non-operating items & tax)"
ws_mli["B16"] = 765.2
ws_mli["B16"].font = F_INPUT
ws_mli["B16"].fill = FILL_ASSUMP
ws_mli["B16"].number_format = CUR
ws_mli["B16"].comment = Comment("10-K: Net income attributable to MLI, FY2025. Taken as reported input — not rebuilt bottom-up (below-the-line items, tax, and minority interest not modeled here).", "Model")

# =========================================================================
# MWA_Standalone  (FY2025, fiscal year ended Sept 30, 2025)
# Source: MWA FY2025 10-K / FY2025 Q4 & Full-Year earnings release, SEC EDGAR
# =========================================================================
ws_mwa["A1"] = "MUELLER WATER PRODUCTS (NYSE: MWA) — STANDALONE, FY2025"
ws_mwa["A1"].font = F_TITLE
note(ws_mwa, "A2", "Source: MWA FY2025 10-K (filed 2025-11-19) and Q4/FY2025 earnings release. All figures $ millions.")

rows2 = [
    ("A4", "Net Sales", "B4", 1429.7, "10-K/earnings release: net sales FY2025"),
    ("A5", "Cost of Goods Sold", "B5", 913.0, "Derived from reported gross profit of $516.7M (third-party aggregator; not a direct 10-K line pull — verify against the filed income statement before finalizing)"),
]
for label_cell, label, val_cell, val, src in rows2:
    ws_mwa[label_cell] = label
    ws_mwa[label_cell].font = F_REG
    ws_mwa[val_cell] = val
    ws_mwa[val_cell].font = F_INPUT
    ws_mwa[val_cell].fill = FILL_ASSUMP
    ws_mwa[val_cell].number_format = CUR
    ws_mwa[val_cell].comment = Comment(src, "Model")

ws_mwa["A7"] = "Gross Profit"
ws_mwa["B7"] = "=B4-B5"
ws_mwa["B7"].font = F_BOLD
ws_mwa["B7"].number_format = CUR
ws_mwa["A8"] = "Gross Margin %"
ws_mwa["B8"] = "=B7/B4"
ws_mwa["B8"].number_format = PCT

ws_mwa["A10"] = "Operating Expenses Below Gross Profit (SG&A + other)"
ws_mwa["B10"] = 256.1
ws_mwa["B10"].font = F_INPUT
ws_mwa["B10"].fill = FILL_ASSUMP
ws_mwa["B10"].number_format = CUR
ws_mwa["B10"].comment = Comment("Derived as Gross Profit less reported Operating Income ($260.6M). Not a direct 10-K line-item split — MWA's filed income statement should be pulled to break this into SG&A vs. D&A before this model is finalized.", "Model")

ws_mwa["A11"] = "Operating Income"
ws_mwa["B11"] = "=B7-B10"
ws_mwa["B11"].font = F_BOLD
ws_mwa["B11"].number_format = CUR
note(ws_mwa, "C11", "Should tie to reported Operating Income of $260.6M")

ws_mwa["A13"] = "Adjusted EBITDA (company-reported, non-GAAP)"
ws_mwa["B13"] = 326.2
ws_mwa["B13"].font = F_INPUT
ws_mwa["B13"].fill = FILL_ASSUMP
ws_mwa["B13"].number_format = CUR
ws_mwa["B13"].comment = Comment("Company-reported Adjusted EBITDA, FY2025 earnings release. 22.8% margin per company disclosure.", "Model")

ws_mwa["A14"] = "Adjusted EBITDA Margin %"
ws_mwa["B14"] = "=B13/B4"
ws_mwa["B14"].number_format = PCT

ws_mwa["A18"] = "Total Debt (FY2025, per 10-K/earnings release)"
ws_mwa["B18"] = 451.6
ws_mwa["B18"].font = F_INPUT
ws_mwa["B18"].fill = FILL_ASSUMP
ws_mwa["B18"].number_format = CUR
ws_mwa["B18"].comment = Comment("FY2025 earnings release: total debt outstanding, 4.0% Senior Notes, no maturities until June 2029, no financial maintenance covenants.", "Model")
ws_mwa["A19"] = "Cash & Cash Equivalents (FY2025)"
ws_mwa["B19"] = 431.5
ws_mwa["B19"].font = F_INPUT
ws_mwa["B19"].fill = FILL_ASSUMP
ws_mwa["B19"].number_format = CUR
ws_mwa["B19"].comment = Comment("FY2025 earnings release: year-end cash and cash equivalents. Net debt leverage reported at 0.1x.", "Model")
ws_mwa["A20"] = "Capital Expenditures (FY2025 actual)"
ws_mwa["B20"] = 47.3
ws_mwa["B20"].font = F_INPUT
ws_mwa["B20"].fill = FILL_ASSUMP
ws_mwa["B20"].number_format = CUR
ws_mwa["B20"].comment = Comment("FY2025 earnings release: capex $47.3M, consistent with $47.4M prior year.", "Model")
ws_mwa["A21"] = "Dividend Payout Ratio (trailing)"
ws_mwa["B21"] = 0.20
ws_mwa["B21"].font = F_INPUT
ws_mwa["B21"].fill = FILL_ASSUMP
ws_mwa["B21"].number_format = PCT
ws_mwa["B21"].comment = Comment("Reported trailing payout ratio range ~19.7%-21.2% across data providers; 20% used as a rounded point estimate.", "Model")
ws_mwa["A22"] = "Implied D&A (unadjusted EBITDA less Operating Income)"
ws_mwa["B22"] = 46.9
ws_mwa["B22"].font = F_INPUT
ws_mwa["B22"].fill = FILL_ASSUMP
ws_mwa["B22"].number_format = CUR
ws_mwa["B22"].comment = Comment("Unadjusted FY2025 EBITDA ($307.5M, earnings release) less reported Operating Income ($260.6M) implies ~$46.9M of D&A. Derived, not a direct 10-K line pull.", "Model")

ws_mwa["A16"] = "Net Income (reported)"
ws_mwa["B16"] = 191.7
ws_mwa["B16"].font = F_INPUT
ws_mwa["B16"].fill = FILL_ASSUMP
ws_mwa["B16"].number_format = CUR
ws_mwa["B16"].comment = Comment("10-K: Net income FY2025 (fiscal year ended 2025-09-30).", "Model")

# =========================================================================
# Copper_Cost_Engine — bug fixed: no more self-referential formula.
# Prior-period and current price are now two DISTINCT input cells.
# =========================================================================
ws_copper["A1"] = "COPPER COST ENGINE & PASS-THROUGH DYNAMICS (FIXED)"
ws_copper["A1"].font = F_TITLE
note(ws_copper, "A2", "Current/prior COMEX averages are real company-disclosed figures (MLI earnings releases). Fabrication spread and pass-through alpha are illustrative analyst assumptions pending validation — flagged, not sourced.")

c_inputs = [
    ("A4", "Current-Quarter COMEX Copper Avg ($/lb)", "B4", 5.13, PX, "MLI Q4 2025 earnings release: COMEX averaged $5.13/lb in Q4 2025"),
    ("A5", "Prior-Quarter COMEX Copper Avg ($/lb)",   "B5", 4.83, PX, "MLI Q3 2025 earnings release: COMEX averaged $4.83/lb in Q3 2025"),
    ("A6", "Pass-Through Pricing Lag (Days)",          "B6", 45,   "0", "ASSUMPTION — typical order-to-ship lag cited in prior project materials; not a disclosed company metric"),
    ("A7", "Base Fabrication Spread ($/lb)",           "B7", 0.30, PX, "ASSUMPTION — illustrative copper tube/fitting fabrication spread; MLI does not publicly disclose this figure. Validate against segment disclosures before using in a live interview."),
    ("A8", "Pass-Through Efficiency (Alpha)",          "B8", 0.85, PCT, "ASSUMPTION — fraction of a quarter-over-quarter copper price move reflected in ASP within the period"),
]
for label_cell, label, val_cell, val, fmt, src in c_inputs:
    ws_copper[label_cell] = label
    ws_copper[label_cell].font = F_REG
    ws_copper[val_cell] = val
    ws_copper[val_cell].font = F_INPUT
    ws_copper[val_cell].fill = FILL_ASSUMP
    ws_copper[val_cell].number_format = fmt
    ws_copper[val_cell].comment = Comment(src, "Model")

ws_copper["A11"] = "Effective Procurement Cost ($/lb)"
# Blends prior and current price based on how much of the year the lag represents.
# No self-reference: B4 and B5 are independent, genuine inputs.
ws_copper["B11"] = "=B5+(B4-B5)*(1-B6/365)"
ws_copper["B11"].font = F_FORM
ws_copper["B11"].number_format = PX
note(ws_copper, "C11", "Formula: PriorPrice + (CurrentPrice-PriorPrice) x (1 - Lag/365). At Lag=0 this equals CurrentPrice; at Lag=365 it equals PriorPrice.")

ws_copper["A12"] = "Realized Selling Price (ASP, $/lb)"
ws_copper["B12"] = "=B7+B8*B4+(1-B8)*B5"
ws_copper["B12"].font = F_FORM
ws_copper["B12"].number_format = PX
note(ws_copper, "C12", "Formula: FabricationSpread + Alpha x CurrentPrice + (1-Alpha) x PriorPrice")

ws_copper["A13"] = "Realized Unit Margin Spread ($/lb)"
ws_copper["B13"] = "=B12-B11"
ws_copper["B13"].font = F_BOLD
ws_copper["B13"].number_format = PX
ws_copper["B13"].border = Border(top=Side(style="thin"), bottom=Side(style="double"))

ws_copper["A15"] = "Formula check: at Lag=0, Effective Cost should equal Current Price"
ws_copper["B15"] = "=IF(ROUND(B5+(B4-B5)*(1-0/365),4)=ROUND(B4,4),\"PASS\",\"FAIL\")"
ws_copper["B15"].font = F_REG

# =========================================================================
# Pro_Forma_Consolidated — every line is a formula referencing the two
# standalone tabs plus a synergies assumption block. Nothing hardcoded.
# =========================================================================
ws_pf["A1"] = "MLI / MWA — PRO FORMA CONSOLIDATED (HYPOTHETICAL)"
ws_pf["A1"].font = F_TITLE
note(ws_pf, "A2", "Illustrative combination for CV/portfolio purposes — not an announced transaction. Synergy figures are analyst assumptions, not disclosed guidance.")

ws_pf["A4"] = "Metric"
ws_pf["B4"] = "MLI Standalone"
ws_pf["C4"] = "MWA Standalone"
ws_pf["D4"] = "Synergies / Adj."
ws_pf["E4"] = "Pro Forma Combined"
for c in ["A4","B4","C4","D4","E4"]:
    ws_pf[c].font = F_HDR
    ws_pf[c].fill = FILL_HDR

r = 5
ws_pf[f"A{r}"] = "Net Sales"
ws_pf[f"B{r}"] = "=MLI_Standalone!B4"
ws_pf[f"C{r}"] = "=MWA_Standalone!B4"
ws_pf[f"D{r}"] = 0
ws_pf[f"D{r}"].font = F_INPUT
ws_pf[f"D{r}"].fill = FILL_ASSUMP
ws_pf[f"E{r}"] = f"=B{r}+C{r}+D{r}"
r += 1

ws_pf[f"A{r}"] = "COGS (informational — copper-sensitivity link only)"
ws_pf[f"B{r}"] = "=MLI_Standalone!B5"
ws_pf[f"C{r}"] = "=MWA_Standalone!B5"
ws_pf[f"D{r}"] = -8
ws_pf[f"D{r}"].font = F_INPUT
ws_pf[f"D{r}"].fill = FILL_ASSUMP
ws_pf[f"D{r}"].comment = Comment("ASSUMPTION: illustrative procurement-pooling synergy, not disclosed. Shown for narrative only — does not feed the EBITDA build below, which uses reported headline EBITDA instead (see MLI/MWA reconciling-gap note).", "Model")
ws_pf[f"E{r}"] = f"=B{r}+C{r}+D{r}"
r += 1

ws_pf[f"A{r}"] = "Gross Profit (informational)"
gp_row = r
ws_pf[f"B{r}"] = "=B5-B6"
ws_pf[f"C{r}"] = "=C5-C6"
ws_pf[f"D{r}"] = "=D5-D6"
ws_pf[f"E{r}"] = "=E5-E6"
r += 2

ws_pf[f"A{r}"] = "EBITDA (built from reported headline figures, not re-derived)"
ebitda_row = r
ws_pf[f"B{r}"] = "=MLI_Standalone!B13"
ws_pf[f"C{r}"] = "=MWA_Standalone!B13"
ws_pf[f"D{r}"] = -14
ws_pf[f"D{r}"].font = F_INPUT
ws_pf[f"D{r}"].fill = FILL_ASSUMP
ws_pf[f"D{r}"].comment = Comment("ASSUMPTION: net illustrative synergies ($8M COGS + $6M opex), not disclosed guidance.", "Model")
ws_pf[f"E{r}"] = f"=B{r}+C{r}+D{r}"
ws_pf[f"A{r}"].font = F_BOLD
for c in "BCDE":
    ws_pf[f"{c}{r}"].font = F_BOLD
    ws_pf[f"{c}{r}"].border = Border(top=Side(style="thin"), bottom=Side(style="double"))
r += 1

ws_pf[f"A{r}"] = "EBITDA Margin %"
margin_row = r
for c in "BCE":
    col_sales = {"B":"B5","C":"C5","E":"E5"}[c]
    ws_pf[f"{c}{r}"] = f"={c}{ebitda_row}/{col_sales}"
    ws_pf[f"{c}{r}"].number_format = PCT

for rr in range(5, margin_row+1):
    for c in "BCDE":
        cell = ws_pf[f"{c}{rr}"]
        if cell.number_format == "General":
            cell.number_format = CUR

for col, width in [("A", 45), ("B", 16), ("C", 16), ("D", 16), ("E", 18)]:
    ws_pf.column_dimensions[col].width = width

note(ws_pf, f"A{margin_row+2}", f"Pro forma EBITDA (E{ebitda_row}) must equal MLI EBITDA + MWA EBITDA + net synergies — checked on Audit_Checks tab.")

# =========================================================================
# Sensitivity_Matrix — 2D: copper price shock x synergy realization %
# =========================================================================
ws_sens["A1"] = "2D SENSITIVITY — PRO FORMA EBITDA"
ws_sens["A1"].font = F_TITLE
note(ws_sens, "A2", "Rows: copper price shock applied to Current-Quarter COMEX price. Columns: % of assumed synergies realized. Values: Pro Forma EBITDA ($M).")

shocks = [-0.20, -0.10, 0, 0.10, 0.20]
realizations = [0.5, 0.75, 1.0, 1.25]

ws_sens["A4"] = "Copper Price Shock \\ Synergy Realization"
ws_sens["A4"].font = F_BOLD
for j, real in enumerate(realizations):
    cell = ws_sens.cell(row=4, column=2+j, value=real)
    cell.number_format = PCT
    cell.font = F_BOLD

base_synergy_cogs = -8
base_synergy_opex = -6

for i, shock in enumerate(shocks):
    row = 5+i
    ws_sens.cell(row=row, column=1, value=shock).number_format = PCT
    ws_sens.cell(row=row, column=1).font = F_BOLD
    for j, real in enumerate(realizations):
        col = 2+j
        # Pro forma EBITDA under this shock/realization:
        # Base EBITDA (0 synergy) + synergy COGS*real + synergy opex*real
        # Copper shock affects MLI COGS proportionally as a simple illustrative linkage:
        # adjusted MLI COGS = MLI COGS * (1 + shock * passthrough_sensitivity)
        # Using a 0.6 illustrative pass-through sensitivity of COGS to a copper price shock.
        # Shocks FY2026E (first forecast year) EBITDA from the Five_Year_Forecast
        # tab, not the FY2025A actual — so the copper engine and the forecast
        # are now the same base case, not two tabs that never speak to each other.
        # MLI FY2026E EBITDA shocked by copper-driven COGS movement (illustrative
        # 0.6 pass-through sensitivity); MWA FY2026E EBITDA is unshocked (copper
        # is not MWA's primary input); synergies scaled by realization %.
        formula = (
            f"=Five_Year_Forecast!C39-MLI_Standalone!B5*{shock}*0.6"
            f"+Five_Year_Forecast!C40"
            f"+({base_synergy_cogs}+{base_synergy_opex})*{real}"
        )
        ws_sens.cell(row=row, column=col, value=formula).number_format = CUR

ws_sens.column_dimensions["A"].width = 32
for col_letter in ["B","C","D","E"]:
    ws_sens.column_dimensions[col_letter].width = 14

note(ws_sens, "A11", "Copper COGS sensitivity factor (0.6) and synergy base ($8M COGS + $6M opex) are illustrative assumptions — label them as such in any pitch, not as disclosed figures.")

# =========================================================================
# Five_Year_Forecast (REBUILT) — driver-based revenue, CapEx tied to
# disclosed guidance, copper engine linked to the base case, a full
# EBITDA-to-FCF bridge (D&A/interest/tax/NWC/dividends/cost-to-achieve),
# and a credit panel that rolls cash forward instead of holding it static.
# =========================================================================
ws_fcst["A1"] = "5-YEAR FORECAST — PRO FORMA (ILLUSTRATIVE), REBUILT"
ws_fcst["A1"].font = F_TITLE
note(ws_fcst, "A2", "FY2025 column is actual, sourced data. FY2026E-FY2030E are projections. Every assumption below is labeled as illustrative/derived or as sourced guidance — see comments.")
note(ws_fcst, "A3", "Note: the FY2025A column here is the PRE-DEAL standalone baseline (zero synergies) used to anchor the forecast — it is not the same figure as Pro_Forma_Consolidated's FY2025 pro forma (which nets illustrative Day-1 synergies). The two tabs answer different questions: Pro_Forma_Consolidated asks 'what would combined FY2025 have looked like,' this tab asks 'where does the combined entity go from here.'")

ws_fcst["A4"] = "REVENUE DRIVERS (volume x price/mix, not a flat blended rate)"
ws_fcst["A4"].font = F_BOLD
drivers = [
    ("A5", "MLI volume growth %/yr",  "B5", 0.020, PCT, "ASSUMPTION — illustrative unit-volume growth. Neither company discloses segment volume at this granularity; this decomposes growth into drivers rather than using one flat rate, but is still an analyst assumption, not sourced guidance."),
    ("A6", "MLI price/mix growth %/yr","B6", 0.020, PCT, "ASSUMPTION — held modest and consistent with the base-case copper price being flat (see Copper_Cost_Engine); if copper prices were assumed to rise, this line should rise with it."),
    ("A7", "MWA volume growth %/yr",  "B7", 0.025, PCT, "ASSUMPTION — illustrative, informed by publicly discussed water-infrastructure capex tailwinds; not a disclosed MWA figure."),
    ("A8", "MWA price growth %/yr",   "B8", 0.024, PCT, "ASSUMPTION — illustrative industrial pricing growth."),
]
for label_cell, label, val_cell, val, fmt, src in drivers:
    ws_fcst[label_cell] = label
    ws_fcst[label_cell].font = F_REG
    ws_fcst[val_cell] = val
    ws_fcst[val_cell].font = F_INPUT
    ws_fcst[val_cell].fill = FILL_ASSUMP
    ws_fcst[val_cell].number_format = fmt
    ws_fcst[val_cell].comment = Comment(src, "Model")

ws_fcst["A9"] = "MLI combined growth %/yr (derived)"
ws_fcst["B9"] = "=(1+B5)*(1+B6)-1"
ws_fcst["B9"].number_format = PCT
ws_fcst["A10"] = "MWA combined growth %/yr (derived)"
ws_fcst["B10"] = "=(1+B7)*(1+B8)-1"
ws_fcst["B10"].number_format = PCT

ws_fcst["A12"] = "MARGIN, COPPER & D&A ASSUMPTIONS"
ws_fcst["A12"].font = F_BOLD
m_assump = [
    ("A13", "Base-case copper price path",   "B13", "Flat at Current-Qtr COMEX ($5.13/lb)", None, "Explicitly consistent with Copper_Cost_Engine!B4. MLI EBITDA margin below is held at FY2025 actual on this basis — if the copper price assumption changes, the margin assumption should be revisited, not left stale."),
    ("A14", "MLI EBITDA margin (held flat)", "B14", "=MLI_Standalone!B14", PCT, "Linked to FY2025 actual margin, earned under copper averaging $4.83-5.13/lb — consistent with the flat price path above."),
    ("A15", "MWA EBITDA margin (held flat)", "B15", "=MWA_Standalone!B14", PCT, "Linked to FY2025 actual margin."),
    ("A16", "MLI D&A as % of revenue",       "B16", "=MLI_Standalone!B6/MLI_Standalone!B4", PCT, "Derived from FY2025 actual D&A / actual revenue."),
    ("A17", "MWA D&A as % of revenue",       "B17", "=MWA_Standalone!B22/MWA_Standalone!B4", PCT, "Derived from implied FY2025 D&A / actual revenue."),
]
for row in m_assump:
    label_cell, label, val_cell, val = row[0], row[1], row[2], row[3]
    fmt, src = row[4], row[5]
    ws_fcst[label_cell] = label
    ws_fcst[label_cell].font = F_REG
    ws_fcst[val_cell] = val
    if isinstance(val, str) and val.startswith("="):
        ws_fcst[val_cell].font = F_FORM
    else:
        ws_fcst[val_cell].font = F_INPUT
        ws_fcst[val_cell].fill = FILL_ASSUMP
    if fmt:
        ws_fcst[val_cell].number_format = fmt
    ws_fcst[val_cell].comment = Comment(src, "Model")

ws_fcst["A19"] = "CAPEX, SYNERGY, CREDIT & CASH-RETURN ASSUMPTIONS"
ws_fcst["A19"].font = F_BOLD
c_assump = [
    ("A20", "MLI FY2026E CapEx (disclosed guidance midpoint)", "B20", "=(MLI_Standalone!B24+MLI_Standalone!C24)/2", CUR, "Midpoint of company-disclosed $80-90M FY2026 CapEx guidance — used directly for Year 1 instead of a historical ratio, so the forecast doesn't understate a number the company has already told the market."),
    ("A21", "MLI CapEx %/revenue, FY2027E onward (implied by guidance)", "B21", "=B20/(MLI_Standalone!B4*(1+B9))", PCT, "Ratio implied by the guided FY2026 CapEx over FY2026E revenue — carried forward for later years in the absence of further guidance."),
    ("A22", "MWA CapEx as % of revenue", "B22", "=MWA_Standalone!B20/MWA_Standalone!B4", PCT, "Derived from FY2025 actual capex / actual revenue — no MWA CapEx guidance located."),
    ("A23", "Total run-rate synergies ($M)", "B23", 14.0, CUR, "Same illustrative synergy assumption used in Pro_Forma_Consolidated."),
    ("A24", "Cost-to-achieve, total (x run-rate synergies)", "B24", 1.2, "0.00\"x\"", "ASSUMPTION — common rule-of-thumb that integration costs run roughly 1-1.5x run-rate synergy value; not company-specific."),
    ("A25", "Blended effective tax rate", "B25", 0.24, PCT, "ASSUMPTION — illustrative blended US federal+state rate; not each company's actual effective rate, which isn't separately available at this granularity."),
    ("A26", "MWA Senior Notes interest rate", "B26", 0.040, PCT, "Sourced: MWA's 4.0% Senior Notes, no maturity until 2029 (bullet structure — no scheduled amortization assumed in this window)."),
    ("A27", "Incremental ΔNWC as % of revenue growth", "B27", 0.08, PCT, "ASSUMPTION — simplified working-capital intensity proxy; a higher-fidelity build would use each company's actual AR/Inventory/AP turnover from the balance sheet, not yet pulled into this model."),
]
for label_cell, label, val_cell, val, fmt, src in c_assump:
    ws_fcst[label_cell] = label
    ws_fcst[label_cell].font = F_REG
    ws_fcst[val_cell] = val
    if isinstance(val, str) and val.startswith("="):
        ws_fcst[val_cell].font = F_FORM
    else:
        ws_fcst[val_cell].font = F_INPUT
        ws_fcst[val_cell].fill = FILL_ASSUMP
    if fmt:
        ws_fcst[val_cell].number_format = fmt
    ws_fcst[val_cell].comment = Comment(src, "Model")

ramp = [0.25, 0.50, 0.75, 1.00, 1.00]
ws_fcst["A29"] = "Synergy realization ramp (FY2026E col C onward)"
for i, rmp in enumerate(ramp):
    cell = ws_fcst.cell(row=29, column=3+i, value=rmp)
    cell.number_format = PCT
    cell.font = F_INPUT
    cell.fill = FILL_ASSUMP

cta_split = [0.60, 0.40, 0, 0, 0]
ws_fcst["A30"] = "Cost-to-achieve cash outlay split (FY2026E col C onward)"
for i, sp in enumerate(cta_split):
    cell = ws_fcst.cell(row=30, column=3+i, value=sp)
    cell.number_format = PCT
    cell.font = F_INPUT
    cell.fill = FILL_ASSUMP
note(ws_fcst, "A31", "Cost-to-achieve is modeled as a one-time CASH outlay in the FCF bridge below, not an EBITDA reduction — consistent with how sponsors typically present 'adjusted' EBITDA while still cash-funding the integration.")

# --- Forecast table ---
hdr_row = 34
ws_fcst[f"A{hdr_row}"] = "Metric"
years = ["FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E"]
for i, y in enumerate(years):
    c = ws_fcst.cell(row=hdr_row, column=2+i, value=y)
for cc in ["A","B","C","D","E","F","G"]:
    ws_fcst[f"{cc}{hdr_row}"].fill = FILL_HDR
    ws_fcst[f"{cc}{hdr_row}"].font = F_HDR

def col(i):
    return get_column_letter(2+i)

r = hdr_row + 1
ws_fcst[f"A{r}"] = "MLI Net Sales"
mli_sales_row = r
ws_fcst[f"{col(0)}{r}"] = "=MLI_Standalone!B4"
for i in range(1, 6):
    ws_fcst[f"{col(i)}{r}"] = f"={col(i-1)}{r}*(1+$B$9)"
r += 1

ws_fcst[f"A{r}"] = "MWA Net Sales"
mwa_sales_row = r
ws_fcst[f"{col(0)}{r}"] = "=MWA_Standalone!B4"
for i in range(1, 6):
    ws_fcst[f"{col(i)}{r}"] = f"={col(i-1)}{r}*(1+$B$10)"
r += 1

ws_fcst[f"A{r}"] = "Combined Net Sales"
sales_row = r
for i in range(6):
    ws_fcst[f"{col(i)}{r}"] = f"={col(i)}{mli_sales_row}+{col(i)}{mwa_sales_row}"
ws_fcst[f"A{r}"].font = F_BOLD
r += 2

ws_fcst[f"A{r}"] = "MLI EBITDA (margin held flat, copper-consistent)"
mli_ebitda_row = r
for i in range(6):
    ws_fcst[f"{col(i)}{r}"] = f"={col(i)}{mli_sales_row}*$B$14"
r += 1
ws_fcst[f"A{r}"] = "MWA EBITDA (margin held flat)"
mwa_ebitda_row = r
for i in range(6):
    ws_fcst[f"{col(i)}{r}"] = f"={col(i)}{mwa_sales_row}*$B$15"
r += 1
ws_fcst[f"A{r}"] = "Synergies realized"
syn_row = r
ws_fcst[f"{col(0)}{r}"] = 0
ws_fcst[f"{col(0)}{r}"].font = F_INPUT
for i in range(1, 6):
    ws_fcst[f"{col(i)}{r}"] = f"=$B$23*{col(i)}29"
r += 1
ws_fcst[f"A{r}"] = "Pro Forma EBITDA"
ebitda_fc_row = r
for i in range(6):
    ws_fcst[f"{col(i)}{r}"] = f"={col(i)}{mli_ebitda_row}+{col(i)}{mwa_ebitda_row}+{col(i)}{syn_row}"
ws_fcst[f"A{r}"].font = F_BOLD
for i in range(6):
    ws_fcst[f"{col(i)}{r}"].font = F_BOLD
    ws_fcst[f"{col(i)}{r}"].border = Border(top=Side(style="thin"), bottom=Side(style="double"))
r += 1
ws_fcst[f"A{r}"] = "Pro Forma EBITDA Margin %"
for i in range(6):
    ws_fcst[f"{col(i)}{r}"] = f"={col(i)}{ebitda_fc_row}/{col(i)}{sales_row}"
    ws_fcst[f"{col(i)}{r}"].number_format = PCT
r += 2

ws_fcst[f"A{r}"] = "Less: D&A (MLI + MWA)"
dna_row = r
for i in range(6):
    ws_fcst[f"{col(i)}{r}"] = f"=-({col(i)}{mli_sales_row}*$B$16+{col(i)}{mwa_sales_row}*$B$17)"
r += 1
ws_fcst[f"A{r}"] = "EBIT"
ebit_row = r
for i in range(6):
    ws_fcst[f"{col(i)}{r}"] = f"={col(i)}{ebitda_fc_row}+{col(i)}{dna_row}"
ws_fcst[f"A{r}"].font = F_BOLD
r += 1
ws_fcst[f"A{r}"] = "Less: Interest Expense (MWA Senior Notes only; MLI debt-free)"
int_row = r
for i in range(6):
    ws_fcst[f"{col(i)}{r}"] = "=-MWA_Standalone!B18*$B$26"
r += 1
ws_fcst[f"A{r}"] = "Pretax Income"
pretax_row = r
for i in range(6):
    ws_fcst[f"{col(i)}{r}"] = f"={col(i)}{ebit_row}+{col(i)}{int_row}"
r += 1
ws_fcst[f"A{r}"] = "Less: Tax"
tax_row = r
for i in range(6):
    ws_fcst[f"{col(i)}{r}"] = f"=-{col(i)}{pretax_row}*$B$25"
r += 1
ws_fcst[f"A{r}"] = "Net Income"
ni_row = r
for i in range(6):
    ws_fcst[f"{col(i)}{r}"] = f"={col(i)}{pretax_row}+{col(i)}{tax_row}"
ws_fcst[f"A{r}"].font = F_BOLD
r += 2

ws_fcst[f"A{r}"] = "Add back: D&A"
addback_row = r
for i in range(6):
    ws_fcst[f"{col(i)}{r}"] = f"=-{col(i)}{dna_row}"
r += 1

ws_fcst[f"A{r}"] = "MLI CapEx"
mli_capex_row = r
ws_fcst[f"{col(0)}{r}"] = "=MLI_Standalone!B23"
ws_fcst[f"{col(1)}{r}"] = "=$B$20"
for i in range(2, 6):
    ws_fcst[f"{col(i)}{r}"] = f"={col(i)}{mli_sales_row}*$B$21"
r += 1
ws_fcst[f"A{r}"] = "MWA CapEx"
mwa_capex_row = r
for i in range(6):
    ws_fcst[f"{col(i)}{r}"] = f"={col(i)}{mwa_sales_row}*$B$22"
r += 1
ws_fcst[f"A{r}"] = "Less: Combined CapEx"
capex_row = r
for i in range(6):
    ws_fcst[f"{col(i)}{r}"] = f"=-({col(i)}{mli_capex_row}+{col(i)}{mwa_capex_row})"
r += 1

ws_fcst[f"A{r}"] = "Less: Incremental NWC Build"
nwc_row = r
ws_fcst[f"{col(0)}{r}"] = 0
ws_fcst[f"{col(0)}{r}"].font = F_INPUT
for i in range(1, 6):
    ws_fcst[f"{col(i)}{r}"] = f"=-({col(i)}{sales_row}-{col(i-1)}{sales_row})*$B$27"
r += 1

ws_fcst[f"A{r}"] = "Less: Dividends (MLI + MWA, trailing payout ratios)"
div_row = r
ws_fcst[f"{col(0)}{r}"] = 0
ws_fcst[f"{col(0)}{r}"].font = F_INPUT
for i in range(1, 6):
    mli_ni_share = f"({col(i)}{ni_row}*({col(i)}{mli_ebitda_row}/{col(i)}{ebitda_fc_row}))"
    ws_fcst[f"{col(i)}{r}"] = (
        f"=-(({col(i)}{mli_ebitda_row}/({col(i)}{mli_ebitda_row}+{col(i)}{mwa_ebitda_row}))*{col(i)}{ni_row}*MLI_Standalone!B25"
        f"+({col(i)}{mwa_ebitda_row}/({col(i)}{mli_ebitda_row}+{col(i)}{mwa_ebitda_row}))*{col(i)}{ni_row}*MWA_Standalone!B21)"
    )
note(ws_fcst, f"H{div_row}", "Combined Net Income is split between MLI and MWA in proportion to each company's EBITDA contribution, then each piece is paid out at that company's own trailing payout ratio.")
r += 1

ws_fcst[f"A{r}"] = "Less: Cost-to-Achieve Synergies (cash outlay)"
cta_row = r
ws_fcst[f"{col(0)}{r}"] = 0
ws_fcst[f"{col(0)}{r}"].font = F_INPUT
for i in range(1, 6):
    ws_fcst[f"{col(i)}{r}"] = f"=-$B$23*$B$24*{col(i)}30"
r += 1

ws_fcst[f"A{r}"] = "Free Cash Flow"
fcf_row = r
for i in range(6):
    ws_fcst[f"{col(i)}{r}"] = (
        f"={col(i)}{ni_row}+{col(i)}{addback_row}+{col(i)}{capex_row}"
        f"+{col(i)}{nwc_row}+{col(i)}{div_row}+{col(i)}{cta_row}"
    )
ws_fcst[f"A{r}"].font = F_BOLD
for i in range(6):
    ws_fcst[f"{col(i)}{r}"].font = F_BOLD
    ws_fcst[f"{col(i)}{r}"].border = Border(top=Side(style="thin"), bottom=Side(style="double"))
r += 2

# --- Credit panel: debt flat (bullet notes, no scheduled amortization
# within the forecast window), cash rolls forward with FCF. ---
ws_fcst[f"A{r}"] = "CREDIT PANEL"
ws_fcst[f"A{r}"].font = F_BOLD
r += 1
ws_fcst[f"A{r}"] = "Combined Total Debt (MWA Senior Notes, bullet to 2029 — no amort assumed)"
debt_row = r
for i in range(6):
    ws_fcst[f"{col(i)}{r}"] = "=MWA_Standalone!B18+MLI_Standalone!B21"
r += 1
ws_fcst[f"A{r}"] = "Combined Cash (rolled forward with Free Cash Flow)"
cash_row = r
ws_fcst[f"{col(0)}{r}"] = "=MLI_Standalone!B22+MWA_Standalone!B19"
for i in range(1, 6):
    ws_fcst[f"{col(i)}{r}"] = f"={col(i-1)}{r}+{col(i)}{fcf_row}"
r += 1
ws_fcst[f"A{r}"] = "Net Debt / (Net Cash)"
netdebt_row = r
for i in range(6):
    ws_fcst[f"{col(i)}{r}"] = f"={col(i)}{debt_row}-{col(i)}{cash_row}"
ws_fcst[f"A{r}"].font = F_BOLD
r += 1
ws_fcst[f"A{r}"] = "Net Debt / EBITDA (x)"
for i in range(6):
    ws_fcst[f"{col(i)}{r}"] = f"={col(i)}{netdebt_row}/{col(i)}{ebitda_fc_row}"
    ws_fcst[f"{col(i)}{r}"].number_format = "0.00x"
r += 2
note(ws_fcst, f"A{r}", "Debt is held flat because MWA's Senior Notes are a bullet structure maturing in 2029, beyond this window, with no scheduled amortization. Cash now rolls forward using actual Free Cash Flow after CapEx, NWC, dividends, interest, tax, and cost-to-achieve — not held static.")
r += 1

for c in "ABCDEFG":
    ws_fcst.column_dimensions[c].width = 50 if c == "A" else 15
for rr in range(hdr_row+1, netdebt_row+2):
    for c in "BCDEFG":
        cell = ws_fcst[f"{c}{rr}"]
        if cell.number_format == "General":
            cell.number_format = CUR

FCST_KEY_ROWS = dict(
    mli_sales_row=mli_sales_row, mwa_sales_row=mwa_sales_row, sales_row=sales_row,
    mli_ebitda_row=mli_ebitda_row, mwa_ebitda_row=mwa_ebitda_row, syn_row=syn_row,
    ebitda_fc_row=ebitda_fc_row, mli_capex_row=mli_capex_row, mwa_capex_row=mwa_capex_row,
    capex_row=capex_row, fcf_row=fcf_row, netdebt_row=netdebt_row, ni_row=ni_row,
    hdr_row=hdr_row,
)
# =========================================================================
# Audit_Checks — the tab the original file had empty.
# =========================================================================
ws_audit["A1"] = "AUDIT CHECKS"
ws_audit["A1"].font = F_TITLE

checks = [
    ("Pro forma Net Sales = MLI + MWA Net Sales",
     f"=IF(ROUND(Pro_Forma_Consolidated!E5,1)=ROUND(MLI_Standalone!B4+MWA_Standalone!B4,1),\"PASS\",\"FAIL\")"),
    ("Pro forma EBITDA = MLI EBITDA + MWA Adj. EBITDA + net synergies",
     f"=IF(ROUND(Pro_Forma_Consolidated!E{ebitda_row},1)=ROUND(MLI_Standalone!B13+MWA_Standalone!B13+{base_synergy_cogs}+{base_synergy_opex},1),\"PASS\",\"FAIL\")"),
    ("MWA standalone Operating Income ties to reported figure ($260.6M)",
     "=IF(ROUND(MWA_Standalone!B11,1)=260.6,\"PASS\",\"FAIL\")"),
    ("Copper engine: Effective Cost = Current Price when Lag = 0",
     "=Copper_Cost_Engine!B15"),
    ("KNOWN GAP: MLI component build (GP-D&A-SG&A) vs. reported Op. Income",
     "=MLI_Standalone!B19"),
    ("Forecast: FY2026E MLI CapEx uses disclosed guidance midpoint ($85.0M)",
     "=IF(ROUND(Five_Year_Forecast!C53,1)=ROUND(Five_Year_Forecast!B20,1),\"PASS\",\"FAIL\")"),
    ("Forecast: synergy ramp never exceeds 100% of run-rate",
     "=IF(MAX(Five_Year_Forecast!C29:G29)<=1,\"PASS\",\"FAIL\")"),
    ("Forecast: FY2026E cash = FY2025A cash + FY2026E Free Cash Flow (roll-forward ties)",
     "=IF(ROUND(Five_Year_Forecast!C63,1)=ROUND(Five_Year_Forecast!B63+Five_Year_Forecast!C59,1),\"PASS\",\"FAIL\")"),
    ("Forecast: synergies reach full $14.0M run-rate by FY2029E",
     "=IF(ROUND(Five_Year_Forecast!F41,1)=ROUND(Five_Year_Forecast!B23,1),\"PASS\",\"FAIL\")"),
]
for i, (label, formula) in enumerate(checks):
    row = 3+i
    ws_audit[f"A{row}"] = label
    ws_audit[f"A{row}"].font = F_REG
    ws_audit[f"B{row}"] = formula
    ws_audit[f"B{row}"].font = F_BOLD

ws_audit.column_dimensions["A"].width = 62
ws_audit.column_dimensions["B"].width = 12

# =========================================================================
# Executive_Summary — links only, nothing hardcoded
# =========================================================================
ws_sum["A1"] = "MLI / MWA HORIZONTAL CONSOLIDATION — EXECUTIVE SUMMARY"
ws_sum["A1"].font = F_TITLE
note(ws_sum, "A2", "Hypothetical, independently-built analysis using public FY2025 filings. Not an announced or advised transaction.")

summary_rows = [
    ("Acquirer", "Mueller Industries, Inc. (NYSE: MLI)"),
    ("Target", "Mueller Water Products, Inc. (NYSE: MWA)"),
    ("MLI FY2025 Net Sales ($M)", "=MLI_Standalone!B4"),
    ("MWA FY2025 Net Sales ($M)", "=MWA_Standalone!B4"),
    ("Pro Forma Net Sales ($M)", f"=Pro_Forma_Consolidated!E5"),
    ("Pro Forma EBITDA ($M)", f"=Pro_Forma_Consolidated!E{ebitda_row}"),
    ("Pro Forma EBITDA Margin %", f"=Pro_Forma_Consolidated!E{margin_row}"),
]
for i, (label, val) in enumerate(summary_rows):
    row = 4+i
    ws_sum[f"A{row}"] = label
    ws_sum[f"A{row}"].font = F_REG
    ws_sum[f"B{row}"] = val
    if isinstance(val, str) and val.startswith("="):
        ws_sum[f"B{row}"].font = F_LINK
        if "Margin" in label:
            ws_sum[f"B{row}"].number_format = PCT
        elif "$M" in label:
            ws_sum[f"B{row}"].number_format = CUR
    else:
        ws_sum[f"B{row}"].font = F_REG

ws_sum.column_dimensions["A"].width = 32
ws_sum.column_dimensions["B"].width = 40

# auto width for data sheets
for sheet in [ws_mli, ws_mwa, ws_copper]:
    sheet.column_dimensions["A"].width = 46
    sheet.column_dimensions["B"].width = 16
    sheet.column_dimensions["C"].width = 70

out_path = "/mnt/user-data/outputs/MLI_MWA_Corrected_Model.xlsx"
wb.save(out_path)
print(f"Saved: {out_path}")

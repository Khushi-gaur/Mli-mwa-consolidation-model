"""
MUELLER INDUSTRIES (NYSE: MLI) & MUELLER WATER PRODUCTS (NYSE: MWA)
HORIZONTAL WATER INFRASTRUCTURE CONSOLIDATION & COPPER BUDGETING MODEL
Automated openpyxl Financial Model Generator
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Initialize Workbook
wb = openpyxl.Workbook()

# Sheet references
ws_summary = wb.active
ws_summary.title = "Executive_Summary"
ws_copper = wb.create_sheet("Copper_Cost_Engine")
ws_mli = wb.create_sheet("MLI_Standalone")
ws_mwa = wb.create_sheet("MWA_Standalone")
ws_proforma = wb.create_sheet("Pro_Forma_Consolidated")
ws_sens = wb.create_sheet("Sensitivity_Matrix")
ws_audit = wb.create_sheet("Audit_Checks")

# Styling conventions
font_title = Font(name="Calibri", size=14, bold=True, color="1E293B")
font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
font_bold = Font(name="Calibri", size=10, bold=True, color="000000")
font_regular = Font(name="Calibri", size=10, color="000000")
font_input = Font(name="Calibri", size=10, bold=True, color="0000FF")    # Blue for Inputs
font_formula = Font(name="Calibri", size=10, color="000000")             # Black for Formulas

fill_header = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
fill_subtotal = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
fill_accent = PatternFill(start_color="EEF2FF", end_color="EEF2FF", fill_type="solid")

thin_border = Border(
    left=Side(style='thin', color='CBD5E1'),
    right=Side(style='thin', color='CBD5E1'),
    top=Side(style='thin', color='CBD5E1'),
    bottom=Side(style='thin', color='CBD5E1')
)
top_thin_bottom_double = Border(
    top=Side(style='thin', color='000000'),
    bottom=Side(style='double', color='000000')
)

# Populate Copper Engine
ws_copper["A1"] = "COPPER COST ENGINE & PASS-THROUGH DYNAMICS"
ws_copper["A1"].font = font_title

inputs = [
    ("A4", "COMEX Copper Benchmark ($/lb)", "B4", 4.50, "$#,##0.00"),
    ("A5", "Pass-Through Pricing Lag (Days)", "B5", 45, "0"),
    ("A6", "Base Fabrication Spread ($/lb)", "B6", 1.82, "$#,##0.00"),
    ("A7", "Pass-Through Efficiency (Alpha)", "B7", 0.98, "0.0%"),
    ("A8", "MLI Annual Base Volume (M lbs)", "B8", 540.0, "#,##0.0"),
]

for label_cell, label, val_cell, val, num_fmt in inputs:
    ws_copper[label_cell] = label
    ws_copper[label_cell].font = font_regular
    ws_copper[val_cell] = val
    ws_copper[val_cell].font = font_input
    ws_copper[val_cell].number_format = num_fmt

ws_copper["A11"] = "Effective Procurement Cost ($/lb)"
ws_copper["B11"] = "=B4-(B4-4.50)*(B5/365)*0.70"
ws_copper["B11"].font = font_formula
ws_copper["B11"].number_format = "$#,##0.000"

ws_copper["A12"] = "Realized Selling Price (ASP $/lb)"
ws_copper["B12"] = "=B6+B7*(B4-(B4-4.50)*(B5/365)*1.10)"
ws_copper["B12"].font = font_formula
ws_copper["B12"].number_format = "$#,##0.000"

ws_copper["A13"] = "Realized Unit Margin Spread ($/lb)"
ws_copper["B13"] = "=B12-B11"
ws_copper["B13"].font = font_bold
ws_copper["B13"].number_format = "$#,##0.000"
ws_copper["B13"].border = top_thin_bottom_double

# Auto-fit columns
for sheet in wb.worksheets:
    sheet.views.sheetView[0].showGridLines = True
    for col in sheet.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        sheet.column_dimensions[col_letter].width = max(max_len + 4, 12)

filename = "Mueller_Consolidated_Copper_Budget_Model.xlsx"
wb.save(filename)
print(f"Successfully generated institutional workbook: {filename}")

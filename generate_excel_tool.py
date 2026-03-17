#!/usr/bin/env python3
"""
Generate an Excel-based Statistical Process Capability Tool.
Creates a multi-tab workbook with formulas, conditional formatting, and charts.
"""
import math
from openpyxl import Workbook
from openpyxl.styles import (
    Font, Alignment, Border, Side, PatternFill, numbers
)
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.worksheet.datavalidation import DataValidation

OUTPUT_FILE = "SPC_Statistical_Calculator.xlsx"

# --- Color palette ---
BLUE = "3B82F6"
DARK_BLUE = "1E3A8A"
GREEN = "10B981"
RED = "EF4444"
ORANGE = "F97316"
GRAY = "6B7280"
LIGHT_GRAY = "F1F5F9"
WHITE = "FFFFFF"
DARK_BG = "1F2937"

# --- Reusable styles ---
HEADER_FONT = Font(name="Calibri", bold=True, size=12, color=WHITE)
HEADER_FILL = PatternFill(start_color=DARK_BLUE, end_color=DARK_BLUE, fill_type="solid")
SUBHEADER_FONT = Font(name="Calibri", bold=True, size=11, color=DARK_BLUE)
SUBHEADER_FILL = PatternFill(start_color=LIGHT_GRAY, end_color=LIGHT_GRAY, fill_type="solid")
INPUT_FILL = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
RESULT_FILL = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
WARN_FILL = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
BAD_FILL = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
GOOD_FILL = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
LABEL_FONT = Font(name="Calibri", size=10, bold=True, color=DARK_BLUE)
VALUE_FONT = Font(name="Calibri", size=11)
TITLE_FONT = Font(name="Calibri", size=16, bold=True, color=DARK_BLUE)
THIN_BORDER = Border(
    left=Side(style="thin", color="D1D5DB"),
    right=Side(style="thin", color="D1D5DB"),
    top=Side(style="thin", color="D1D5DB"),
    bottom=Side(style="thin", color="D1D5DB"),
)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def apply_header_row(ws, row, cols, texts):
    """Apply header style to a row."""
    for i, (col, text) in enumerate(zip(cols, texts)):
        cell = ws.cell(row=row, column=col, value=text)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER
        cell.border = THIN_BORDER


def apply_label_value_pair(ws, row, label_col, label, value_col, value=None,
                           formula=None, is_input=False, num_format=None):
    """Write a label-value pair with styling."""
    lc = ws.cell(row=row, column=label_col, value=label)
    lc.font = LABEL_FONT
    lc.alignment = LEFT
    lc.border = THIN_BORDER

    vc = ws.cell(row=row, column=value_col)
    if formula:
        vc.value = formula
    elif value is not None:
        vc.value = value
    vc.font = VALUE_FONT
    vc.alignment = CENTER
    vc.border = THIN_BORDER
    vc.fill = INPUT_FILL if is_input else RESULT_FILL
    if num_format:
        vc.number_format = num_format
    return vc


def create_analysis_sheet(wb):
    """Sheet 1: Analysis — Input parameters and calculated results."""
    ws = wb.active
    ws.title = "Analysis"
    ws.sheet_properties.tabColor = BLUE

    # Column widths
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 4
    ws.column_dimensions["E"].width = 28
    ws.column_dimensions["F"].width = 16
    ws.column_dimensions["G"].width = 4
    ws.column_dimensions["H"].width = 28
    ws.column_dimensions["I"].width = 16

    # --- Title ---
    ws.merge_cells("B1:I1")
    title_cell = ws.cell(row=1, column=2, value="Statistical Process Capability & Optimization Tool")
    title_cell.font = TITLE_FONT
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("B2:I2")
    sub = ws.cell(row=2, column=2, value="Automotive Dimensional Capability Analysis — Single Characteristic")
    sub.font = Font(name="Calibri", size=10, italic=True, color=GRAY)
    sub.alignment = Alignment(horizontal="center")

    # ========== SECTION 1: SPECIFICATIONS (Inputs) ==========
    row = 4
    apply_header_row(ws, row, [2, 3], ["I. SPECIFICATIONS", "Value"])

    row = 5
    apply_label_value_pair(ws, row, 2, "Measurement Name", 3, value="Diameter A", is_input=True)
    row = 6
    apply_label_value_pair(ws, row, 2, "Target Mean (Tₘ)", 3, value=10.00, is_input=True, num_format="0.000")
    row = 7
    apply_label_value_pair(ws, row, 2, "Lower Spec Limit (LSL)", 3, value=9.90, is_input=True, num_format="0.000")
    row = 8
    apply_label_value_pair(ws, row, 2, "Upper Spec Limit (USL)", 3, value=10.10, is_input=True, num_format="0.000")

    # ========== SECTION 2: PROCESS DATA (Inputs) ==========
    row = 10
    apply_header_row(ws, row, [2, 3], ["II. PROCESS DATA", "Value"])

    row = 11
    apply_label_value_pair(ws, row, 2, "Measured Mean (x̄)", 3, value=10.00, is_input=True, num_format="0.000")
    row = 12
    apply_label_value_pair(ws, row, 2, "Standard Deviation (σ)", 3, value=0.015, is_input=True, num_format="0.00000")
    row = 13
    apply_label_value_pair(ws, row, 2, "Sample Size (n)", 3, value=30, is_input=True, num_format="0")
    row = 14
    apply_label_value_pair(ws, row, 2, "Target Capability Index", 3, value=1.67, is_input=True, num_format="0.00")
    row = 15
    apply_label_value_pair(ws, row, 2, "Confidence Level (%)", 3, value=95.0, is_input=True, num_format="0.0")

    # Data Source note
    row = 16
    note = ws.cell(row=row, column=2, value="💡 Or use the 'Data' sheet to enter part-by-part values. Mean & σ auto-calculate there.")
    note.font = Font(name="Calibri", size=9, italic=True, color=GRAY)
    ws.merge_cells(f"B{row}:C{row}")

    # ========== SECTION 3: CALCULATED RESULTS ==========
    row = 4
    apply_header_row(ws, row, [5, 6], ["III. CALCULATED RESULTS", "Value"])

    # --- Named cell references for formulas ---
    # C6=Tm, C7=LSL, C8=USL, C11=x_bar, C12=sigma, C13=n, C14=target_index, C15=CL
    tm, lsl, usl = "C6", "C7", "C8"
    xbar, sigma, n_samp = "C11", "C12", "C13"
    target_idx, cl = "C14", "C15"

    # Tolerance (T = USL - LSL)
    row = 5
    apply_label_value_pair(ws, row, 5, "Tolerance (T = USL − LSL)", 6,
                           formula=f"={usl}-{lsl}", num_format="0.000")

    # 6σ Spread
    row = 6
    apply_label_value_pair(ws, row, 5, "6σ Spread", 6,
                           formula=f"=6*{sigma}", num_format="0.000")

    # 8σ Spread
    row = 7
    apply_label_value_pair(ws, row, 5, "8σ Spread", 6,
                           formula=f"=8*{sigma}", num_format="0.000")

    # Cp
    row = 8
    apply_label_value_pair(ws, row, 5, "Cp (Potential Capability)", 6,
                           formula=f'=IF({sigma}=0,"∞",({usl}-{lsl})/(6*{sigma}))',
                           num_format="0.000")

    # Cpk
    row = 9
    apply_label_value_pair(ws, row, 5, "Cpk (Actual Capability)", 6,
                           formula=f'=IF({sigma}=0,"∞",MIN(({usl}-{xbar})/(3*{sigma}),({xbar}-{lsl})/(3*{sigma})))',
                           num_format="0.000")

    # Required Shift
    row = 10
    apply_label_value_pair(ws, row, 5, "Required Shift (Δ = Tₘ − x̄)", 6,
                           formula=f"={tm}-{xbar}", num_format="0.000")

    # Shift Direction
    row = 11
    apply_label_value_pair(ws, row, 5, "Shift Direction", 6,
                           formula=f'=IF({tm}-{xbar}=0,"Centered",IF({tm}-{xbar}>0,"Shift UP (+)","Shift DOWN (−)"))')

    # Required Tolerance
    row = 12
    apply_label_value_pair(ws, row, 5, "Required Tolerance (for target index)", 6,
                           formula=f"=IF({sigma}=0,0,{target_idx}*6*{sigma})",
                           num_format="0.000")

    # x̄ ± 3σ
    row = 13
    apply_label_value_pair(ws, row, 5, "x̄ − 3σ", 6,
                           formula=f"={xbar}-3*{sigma}", num_format="0.000")
    row = 14
    apply_label_value_pair(ws, row, 5, "x̄ + 3σ", 6,
                           formula=f"={xbar}+3*{sigma}", num_format="0.000")

    # x̄ ± 4σ
    row = 15
    apply_label_value_pair(ws, row, 5, "x̄ − 4σ", 6,
                           formula=f"={xbar}-4*{sigma}", num_format="0.000")
    row = 16
    apply_label_value_pair(ws, row, 5, "x̄ + 4σ", 6,
                           formula=f"={xbar}+4*{sigma}", num_format="0.000")

    # ========== SECTION 4: PROBABILITY & DEFECT ANALYSIS ==========
    row = 18
    apply_header_row(ws, row, [5, 6], ["IV. PROBABILITY & DEFECTS", "Value"])

    # P(x > USL) — uses NORM.S.DIST
    row = 19
    apply_label_value_pair(ws, row, 5, "P(x > USL)", 6,
                           formula=f'=IF({sigma}=0,IF({xbar}>{usl},1,0),1-NORM.DIST({usl},{xbar},{sigma},TRUE))',
                           num_format="0.0000%")

    # P(x < LSL)
    row = 20
    apply_label_value_pair(ws, row, 5, "P(x < LSL)", 6,
                           formula=f'=IF({sigma}=0,IF({xbar}<{lsl},1,0),NORM.DIST({lsl},{xbar},{sigma},TRUE))',
                           num_format="0.0000%")

    # P(x < Tₘ)
    row = 21
    apply_label_value_pair(ws, row, 5, "P(x < Tₘ)", 6,
                           formula=f'=IF({sigma}=0,IF({xbar}<{tm},1,0),NORM.DIST({tm},{xbar},{sigma},TRUE))',
                           num_format="0.00%")

    # PPM Above USL
    row = 22
    apply_label_value_pair(ws, row, 5, "PPM Above USL", 6,
                           formula=f"=F19*1000000", num_format="#,##0.0")

    # PPM Below LSL
    row = 23
    apply_label_value_pair(ws, row, 5, "PPM Below LSL", 6,
                           formula=f"=F20*1000000", num_format="#,##0.0")

    # Total PPM
    row = 24
    apply_label_value_pair(ws, row, 5, "Total PPM (Defect Rate)", 6,
                           formula=f"=F22+F23", num_format="#,##0.0")
    ws["F24"].fill = WARN_FILL

    # ========== SECTION 5: HYPOTHESIS TEST ==========
    row = 18
    apply_header_row(ws, row, [8, 9], ["V. HYPOTHESIS TEST (μ vs Tₘ)", "Value"])

    # Standard Error
    row = 19
    apply_label_value_pair(ws, row, 8, "Standard Error (SE)", 9,
                           formula=f"=IF({sigma}=0,0,{sigma}/SQRT({n_samp}))",
                           num_format="0.00000")

    # Z-statistic
    row = 20
    apply_label_value_pair(ws, row, 8, "Z-statistic", 9,
                           formula=f'=IF(I19=0,IF({xbar}={tm},0,999),({xbar}-{tm})/I19)',
                           num_format="0.000")

    # p-value (Two-Sided)
    row = 21
    apply_label_value_pair(ws, row, 8, "p-value (Two-Sided)", 9,
                           formula=f"=2*(1-NORM.S.DIST(ABS(I20),TRUE))",
                           num_format="0.0000")

    # Alpha
    row = 22
    apply_label_value_pair(ws, row, 8, "Alpha (α)", 9,
                           formula=f"=1-{cl}/100",
                           num_format="0.00")

    # Decision
    row = 23
    apply_label_value_pair(ws, row, 8, "Decision", 9,
                           formula=f'=IF(I21<I22,"Reject H₀ — Mean has shifted","Fail to Reject H₀ — No significant shift")')

    # CI Lower
    row = 24
    apply_label_value_pair(ws, row, 8, "CI Lower Bound", 9,
                           formula=f"={xbar}-NORM.S.INV(1-I22/2)*I19",
                           num_format="0.0000")

    # CI Upper
    row = 25
    apply_label_value_pair(ws, row, 8, "CI Upper Bound", 9,
                           formula=f"={xbar}+NORM.S.INV(1-I22/2)*I19",
                           num_format="0.0000")

    # ========== SECTION 6: VERDICT ==========
    row = 26
    ws.merge_cells("B26:C26")
    apply_header_row(ws, row, [2, 5, 6], ["", "VI. OVERALL VERDICT", ""])

    row = 27
    vc = ws.cell(row=row, column=5,
                 value='=IF(F9="∞","✅ GOOD — Perfect capability (σ=0)",'
                       'IF(F9>=C14,"✅ GOOD — Process is capable",'
                       'IF(F9>=1,"⚠️ MARGINAL — Acceptable but below target",'
                       '"❌ ACTION REQUIRED — Process not capable")))')
    vc.font = Font(name="Calibri", size=12, bold=True)
    vc.alignment = LEFT
    ws.merge_cells("E27:I27")

    # Conditional formatting for Cpk
    ws.conditional_formatting.add("F9",
        CellIsRule(operator="greaterThanOrEqual", formula=["1.67"],
                   fill=GOOD_FILL, font=Font(color="047857", bold=True)))
    ws.conditional_formatting.add("F9",
        CellIsRule(operator="between", formula=["1", "1.669"],
                   fill=WARN_FILL, font=Font(color="92400E", bold=True)))
    ws.conditional_formatting.add("F9",
        CellIsRule(operator="lessThan", formula=["1"],
                   fill=BAD_FILL, font=Font(color="991B1B", bold=True)))

    # Conditional formatting for Cp
    ws.conditional_formatting.add("F8",
        CellIsRule(operator="greaterThanOrEqual", formula=["1.67"],
                   fill=GOOD_FILL, font=Font(color="047857", bold=True)))
    ws.conditional_formatting.add("F8",
        CellIsRule(operator="lessThan", formula=["1"],
                   fill=BAD_FILL, font=Font(color="991B1B", bold=True)))

    # Print setup
    ws.sheet_view.showGridLines = False
    ws.print_area = "A1:I27"


def create_data_sheet(wb):
    """Sheet 2: Data Worksheet — Part-by-part entry with auto-calculations."""
    ws = wb.create_sheet("Data")
    ws.sheet_properties.tabColor = GREEN

    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 6
    ws.column_dimensions["C"].width = 24
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 4
    ws.column_dimensions["F"].width = 24
    ws.column_dimensions["G"].width = 18

    # Title
    ws.merge_cells("B1:G1")
    t = ws.cell(row=1, column=2, value="Data Worksheet — Part-by-Part Entry")
    t.font = TITLE_FONT
    t.alignment = Alignment(horizontal="center")

    # Instructions
    ws.merge_cells("B2:G2")
    inst = ws.cell(row=2, column=2,
                   value="Enter DMC/Serial Number and measured Value for each part below. Statistics auto-calculate on the right.")
    inst.font = Font(name="Calibri", size=9, italic=True, color=GRAY)
    inst.alignment = Alignment(horizontal="center")

    # Data table headers
    row = 4
    apply_header_row(ws, row, [2, 3, 4], ["#", "DMC / Serial Number", "Value"])

    # Data entry rows (50 rows)
    MAX_ROWS = 50
    for i in range(1, MAX_ROWS + 1):
        r = row + i
        ws.cell(row=r, column=2, value=i).font = Font(size=9, color=GRAY)
        ws.cell(row=r, column=2).alignment = CENTER
        ws.cell(row=r, column=2).border = THIN_BORDER
        ws.cell(row=r, column=3).border = THIN_BORDER
        ws.cell(row=r, column=3).fill = INPUT_FILL
        ws.cell(row=r, column=4).border = THIN_BORDER
        ws.cell(row=r, column=4).fill = INPUT_FILL
        ws.cell(row=r, column=4).number_format = "0.000"

    # --- Auto-calculated statistics panel ---
    row = 4
    apply_header_row(ws, row, [6, 7], ["STATISTICS (Auto-Calculated)", "Value"])

    data_range = "D5:D54"

    stats = [
        ("Count (n)", f"=COUNTA({data_range})", "0"),
        ("Mean (x̄)", f"=IF(G5>=2,AVERAGE({data_range}),\"\")", "0.00000"),
        ("Std Dev (σ)", f"=IF(G5>=2,STDEV.S({data_range}),\"\")", "0.00000"),
        ("Min", f"=IF(G5>=1,MIN({data_range}),\"\")", "0.000"),
        ("Max", f"=IF(G5>=1,MAX({data_range}),\"\")", "0.000"),
        ("Range", f"=IF(G5>=2,G9-G8,\"\")", "0.000"),
        ("Median", f"=IF(G5>=1,MEDIAN({data_range}),\"\")", "0.000"),
        ("Q1 (25th %ile)", f"=IF(G5>=4,PERCENTILE({data_range},0.25),\"\")", "0.000"),
        ("Q3 (75th %ile)", f"=IF(G5>=4,PERCENTILE({data_range},0.75),\"\")", "0.000"),
        ("IQR", f"=IF(G5>=4,G13-G12,\"\")", "0.000"),
    ]

    for i, (label, formula, fmt) in enumerate(stats):
        r = 5 + i
        apply_label_value_pair(ws, r, 6, label, 7, formula=formula, num_format=fmt)

    # Quick-link instruction
    row = 16
    ws.merge_cells("F16:G16")
    tip = ws.cell(row=row, column=6,
                  value="💡 Copy G5→C13, G6→C11, G7→C12 in the Analysis sheet to use this data.")
    tip.font = Font(name="Calibri", size=9, italic=True, color=ORANGE)

    # Alternating row colors for data entry
    for i in range(1, MAX_ROWS + 1):
        r = 4 + i
        if i % 2 == 0:
            for col in [2, 3, 4]:
                c = ws.cell(row=r, column=col)
                if c.fill == INPUT_FILL:
                    c.fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")

    ws.sheet_view.showGridLines = False


def create_charts_sheet(wb):
    """Sheet 3: Charts — Histogram and capability visualization using helper data."""
    ws = wb.create_sheet("Charts")
    ws.sheet_properties.tabColor = ORANGE

    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 14

    # Title
    ws.merge_cells("B1:K1")
    t = ws.cell(row=1, column=2, value="Visualization — Capability Charts")
    t.font = TITLE_FONT
    t.alignment = Alignment(horizontal="center")

    # --- Helper data for a bell curve (pre-calculated standard normal z-values) ---
    # We'll generate a bell curve using z = -4 to +4 in steps of 0.5
    # The chart plots: x = x_bar + z*sigma, y = normal PDF

    ws.cell(row=3, column=2, value="Bell Curve Helper Data").font = SUBHEADER_FONT

    ws.cell(row=4, column=2, value="z").font = HEADER_FONT
    ws.cell(row=4, column=2).fill = HEADER_FILL
    ws.cell(row=4, column=3, value="x Value").font = HEADER_FONT
    ws.cell(row=4, column=3).fill = HEADER_FILL
    ws.cell(row=4, column=4, value="Density").font = HEADER_FONT
    ws.cell(row=4, column=4).fill = HEADER_FILL

    z_values = [i * 0.5 for i in range(-8, 9)]  # -4 to +4 in 0.5 steps
    xbar_ref = "Analysis!C11"
    sigma_ref = "Analysis!C12"

    for i, z in enumerate(z_values):
        r = 5 + i
        ws.cell(row=r, column=2, value=z).number_format = "0.0"
        # x = x_bar + z * sigma
        ws.cell(row=r, column=3,
                value=f"={xbar_ref}+B{r}*{sigma_ref}").number_format = "0.000"
        # y = (1/(sigma*SQRT(2*PI)))*EXP(-0.5*z^2)
        ws.cell(row=r, column=4,
                value=f'=IF({sigma_ref}>0,(1/({sigma_ref}*SQRT(2*PI())))*EXP(-0.5*B{r}^2),0)').number_format = "0.0000"

    last_data_row = 5 + len(z_values) - 1

    # --- Bell Curve Chart ---
    chart1 = LineChart()
    chart1.title = "Process Distribution (Normal Curve)"
    chart1.y_axis.title = "Density"
    chart1.x_axis.title = "Measurement Value"
    chart1.style = 10
    chart1.width = 22
    chart1.height = 14

    x_data = Reference(ws, min_col=3, min_row=4, max_row=last_data_row)
    y_data = Reference(ws, min_col=4, min_row=4, max_row=last_data_row)

    chart1.add_data(y_data, titles_from_data=True)
    chart1.set_categories(x_data)
    chart1.series[0].graphicalProperties.line.width = 25000  # ~2pt

    ws.add_chart(chart1, "F3")

    # --- Specification summary below ---
    summary_row = last_data_row + 2
    ws.cell(row=summary_row, column=2, value="Specification Lines Reference").font = SUBHEADER_FONT
    ws.cell(row=summary_row + 1, column=2, value="LSL").font = LABEL_FONT
    ws.cell(row=summary_row + 1, column=3, value=f"=Analysis!C7").number_format = "0.000"
    ws.cell(row=summary_row + 2, column=2, value="USL").font = LABEL_FONT
    ws.cell(row=summary_row + 2, column=3, value=f"=Analysis!C8").number_format = "0.000"
    ws.cell(row=summary_row + 3, column=2, value="Target (Tₘ)").font = LABEL_FONT
    ws.cell(row=summary_row + 3, column=3, value=f"=Analysis!C6").number_format = "0.000"
    ws.cell(row=summary_row + 4, column=2, value="Mean (x̄)").font = LABEL_FONT
    ws.cell(row=summary_row + 4, column=3, value=f"=Analysis!C11").number_format = "0.000"

    # --- Capability Bar Chart ---
    ws.cell(row=3, column=14, value="Metric").font = HEADER_FONT
    ws.cell(row=3, column=14).fill = HEADER_FILL
    ws.cell(row=3, column=15, value="Value").font = HEADER_FONT
    ws.cell(row=3, column=15).fill = HEADER_FILL
    ws.cell(row=3, column=16, value="Target").font = HEADER_FONT
    ws.cell(row=3, column=16).fill = HEADER_FILL

    ws.cell(row=4, column=14, value="Cp")
    ws.cell(row=4, column=15, value="=Analysis!F8").number_format = "0.00"
    ws.cell(row=4, column=16, value=f"=Analysis!C14").number_format = "0.00"
    ws.cell(row=5, column=14, value="Cpk")
    ws.cell(row=5, column=15, value="=Analysis!F9").number_format = "0.00"
    ws.cell(row=5, column=16, value=f"=Analysis!C14").number_format = "0.00"

    chart2 = BarChart()
    chart2.title = "Capability Index vs Target"
    chart2.type = "col"
    chart2.style = 10
    chart2.width = 14
    chart2.height = 12
    chart2.y_axis.title = "Index Value"

    cats = Reference(ws, min_col=14, min_row=4, max_row=5)
    vals = Reference(ws, min_col=15, min_row=3, max_row=5)
    tgts = Reference(ws, min_col=16, min_row=3, max_row=5)
    chart2.add_data(vals, titles_from_data=True)
    chart2.add_data(tgts, titles_from_data=True)
    chart2.set_categories(cats)

    ws.add_chart(chart2, "F20")

    ws.sheet_view.showGridLines = False


def create_history_sheet(wb):
    """Sheet 4: History — Template for recording analysis runs."""
    ws = wb.create_sheet("History")
    ws.sheet_properties.tabColor = "8B5CF6"  # Purple

    ws.column_dimensions["A"].width = 2

    # Title
    ws.merge_cells("B1:K1")
    t = ws.cell(row=1, column=2, value="Analysis History Log")
    t.font = TITLE_FONT
    t.alignment = Alignment(horizontal="center")

    ws.merge_cells("B2:K2")
    inst = ws.cell(row=2, column=2,
                   value="Record each analysis run below. Copy values from the Analysis sheet after each run.")
    inst.font = Font(name="Calibri", size=9, italic=True, color=GRAY)
    inst.alignment = Alignment(horizontal="center")

    # Headers
    headers = [
        "Date", "Characteristic", "Tₘ", "LSL", "USL", "x̄", "σ", "n",
        "Cp", "Cpk", "PPM Total", "Shift (Δ)", "Verdict"
    ]
    col_widths = [12, 18, 10, 10, 10, 12, 12, 8, 10, 10, 12, 12, 30]

    for i, (h, w) in enumerate(zip(headers, col_widths)):
        col = i + 2
        ws.column_dimensions[get_column_letter(col)].width = w
        cell = ws.cell(row=4, column=col, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER
        cell.border = THIN_BORDER

    # Pre-fill first row with formulas linking to Analysis sheet
    r = 5
    formulas = [
        ("=TODAY()", "YYYY-MM-DD"),
        ("=Analysis!C5", None),
        ("=Analysis!C6", "0.000"),
        ("=Analysis!C7", "0.000"),
        ("=Analysis!C8", "0.000"),
        ("=Analysis!C11", "0.000"),
        ("=Analysis!C12", "0.00000"),
        ("=Analysis!C13", "0"),
        ("=Analysis!F8", "0.000"),
        ("=Analysis!F9", "0.000"),
        ("=Analysis!F24", "#,##0.0"),
        ("=Analysis!F10", "0.000"),
        ("=Analysis!E27", None),
    ]

    for i, (formula, fmt) in enumerate(formulas):
        col = i + 2
        cell = ws.cell(row=r, column=col, value=formula)
        cell.border = THIN_BORDER
        cell.alignment = CENTER
        if fmt:
            cell.number_format = fmt
        cell.fill = RESULT_FILL

    # Empty rows for future entries
    for row_idx in range(6, 56):
        for col in range(2, 15):
            cell = ws.cell(row=row_idx, column=col)
            cell.border = THIN_BORDER
            cell.fill = INPUT_FILL if row_idx % 2 == 0 else PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")

    # Conditional formatting for Cpk column
    ws.conditional_formatting.add("K5:K55",
        CellIsRule(operator="greaterThanOrEqual", formula=["1.67"],
                   fill=GOOD_FILL, font=Font(color="047857", bold=True)))
    ws.conditional_formatting.add("K5:K55",
        CellIsRule(operator="between", formula=["1", "1.669"],
                   fill=WARN_FILL, font=Font(color="92400E", bold=True)))
    ws.conditional_formatting.add("K5:K55",
        CellIsRule(operator="lessThan", formula=["1"],
                   fill=BAD_FILL, font=Font(color="991B1B", bold=True)))

    ws.sheet_view.showGridLines = False


def create_reference_sheet(wb):
    """Sheet 5: Reference — Formula definitions and usage guide."""
    ws = wb.create_sheet("Reference")
    ws.sheet_properties.tabColor = GRAY

    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 80

    # Title
    ws.merge_cells("B1:B1")
    t = ws.cell(row=1, column=2, value="Reference Guide — Formulas & Definitions")
    t.font = TITLE_FONT

    content = [
        ("Core Capability Formulas", True),
        ("• Cp = (USL − LSL) / 6σ  — Potential capability if process is perfectly centered", False),
        ("• Cpk = min[(USL − x̄) / 3σ,  (x̄ − LSL) / 3σ]  — Actual capability with centering error", False),
        ("• Required Shift (Δ) = Tₘ − x̄  — How far to adjust the process mean", False),
        ("• Required Tolerance = Target Index × 6σ  — Minimum tolerance band needed", False),
        ("", False),
        ("Capability Index Interpretation", True),
        ("• Cpk ≥ 1.67  →  ✅ GOOD — Process is capable (automotive standard)", False),
        ("• 1.33 ≤ Cpk < 1.67  →  ⚠️ ACCEPTABLE — Meets minimum but below target", False),
        ("• 1.00 ≤ Cpk < 1.33  →  ⚠️ MARGINAL — High risk, improvement needed", False),
        ("• Cpk < 1.00  →  ❌ NOT CAPABLE — Process produces significant defects", False),
        ("", False),
        ("Probability & PPM", True),
        ("• P(x > USL) — Probability a part exceeds upper spec limit", False),
        ("• P(x < LSL) — Probability a part falls below lower spec limit", False),
        ("• PPM = Probability × 1,000,000 — Parts Per Million defective", False),
        ("", False),
        ("Hypothesis Testing", True),
        ("• H₀: μ = Tₘ (Null hypothesis — process is on target)", False),
        ("• Z = (x̄ − Tₘ) / (σ / √n) — Test statistic", False),
        ("• p-value < α → Reject H₀ → Significant evidence mean has shifted", False),
        ("• p-value ≥ α → Fail to Reject H₀ → No significant shift detected", False),
        ("", False),
        ("Spread Metrics", True),
        ("• 6σ Spread — Contains ~99.73% of process output (±3σ)", False),
        ("• 8σ Spread — Contains ~99.9937% of process output (±4σ)", False),
        ("", False),
        ("How to Use This Workbook", True),
        ("1. Enter specifications (Tₘ, LSL, USL) in the Analysis sheet", False),
        ("2. Enter part data in the Data sheet OR enter x̄ and σ manually in Analysis", False),
        ("3. If using Data sheet, copy n→C13, x̄→C11, σ→C12 in the Analysis sheet", False),
        ("4. Read results in the Analysis sheet — Cp, Cpk, PPM, shift, verdict", False),
        ("5. View charts in the Charts sheet", False),
        ("6. Copy results to History sheet to track runs over time", False),
    ]

    row = 3
    for text, is_header in content:
        cell = ws.cell(row=row, column=2, value=text)
        if is_header:
            cell.font = Font(name="Calibri", size=12, bold=True, color=DARK_BLUE)
            cell.fill = SUBHEADER_FILL
        else:
            cell.font = Font(name="Calibri", size=10)
        cell.alignment = LEFT
        row += 1

    ws.sheet_view.showGridLines = False


def main():
    wb = Workbook()

    create_analysis_sheet(wb)
    create_data_sheet(wb)
    create_charts_sheet(wb)
    create_history_sheet(wb)
    create_reference_sheet(wb)

    wb.save(OUTPUT_FILE)
    print(f"✅ Excel tool generated: {OUTPUT_FILE}")
    print(f"   Sheets: Analysis | Data | Charts | History | Reference")


if __name__ == "__main__":
    main()

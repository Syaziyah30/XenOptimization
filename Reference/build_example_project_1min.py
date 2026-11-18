"""Rebuild ExampleProject.xlsx with the client's real MT Chiller 1 data (Hist1-5)."""
import sys
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.chart import ScatterChart, LineChart, BarChart, Reference, Series
from openpyxl.formatting.rule import CellIsRule, FormulaRule

SRC = r"D:\XenOptimization\20261016_benny - Copy.xlsx"
OUT = sys.argv[1]
RAWPKL = sys.argv[2]
P = "C10A1_1_C101C_MTCHW_CH01"

# ---------- data ----------
raw = pd.read_pickle(RAWPKL)
ren = {P + s: n for s, n in [("_CL1.PV", "CL1"), ("_EF.PV", "EF"), ("_EF.PV2", "EF2"), ("_FIT1.PV", "FIT1"),
                              ("_FIT2.PV", "FIT2"), ("_HR1.PV", "HR1"), ("_TIT1.PV", "TIT1"), ("_TIT2.PV", "TIT2"),
                              ("_TIT3.PV", "TIT3"), ("_TIT4.PV", "TIT4")]}
d = raw[["DateTime", "Period"] + list(ren)].rename(columns=ren)
d["DateTime"] = pd.to_datetime(d["DateTime"])
cols = ["CL1", "EF", "EF2", "FIT1", "FIT2", "HR1", "TIT1", "TIT2", "TIT3", "TIT4"]
a = d.sort_values("DateTime").reset_index(drop=True)
a["slot"] = a["DateTime"]
N = len(a)
_tr = a[a.Period != "Hist5"]
RT_LO, RT_HI = (_tr.CL1 / 3.517).min(), (_tr.CL1 / 3.517).max()
T1_LO, T1_HI = _tr.TIT1.min(), _tr.TIT1.max()
RANGE_TXT = f"RT {RT_LO:,.0f}–{RT_HI:,.0f}; TIT1 {T1_LO:.1f}–{T1_HI:.1f} °C"
LAST = N + 1  # last data row on Raw_Data / Cleaning / Clean_Data

# live stand-in: 1-minute readings, 2 Feb 2026 08:00-15:30 every 30 min
live = d[(d.Period == "Hist5") & (d.DateTime.dt.date == pd.Timestamp("2026-02-02").date())
         & (d.DateTime.dt.minute.isin([0, 30])) & (d.DateTime.dt.hour.between(8, 15))].head(16)

# ---------- styles ----------
NAVY, BLUE = "14213D", "1F6FB2"
F_TITLE = Font(name="Arial", size=16, bold=True, color=NAVY)
F_H2 = Font(name="Arial", size=12, bold=True, color=BLUE)
F_HDR = Font(name="Arial", size=10, bold=True, color="FFFFFF")
F_B = Font(name="Arial", size=10, bold=True)
F_N = Font(name="Arial", size=10)
F_I = Font(name="Arial", size=9, italic=True, color="595959")
F_IN = Font(name="Arial", size=10, color="0000FF")
FILL_HDR = PatternFill("solid", fgColor=NAVY)
FILL_IN = PatternFill("solid", fgColor="FFF2CC")
FILL_KEY = PatternFill("solid", fgColor="DDEBF7")
FILL_REC = PatternFill("solid", fgColor="FCE4D6")
FILL_SEC = PatternFill("solid", fgColor="E7E6E6")
WRAP = Alignment(wrap_text=True, vertical="top")
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = Workbook()
wb.remove(wb.active)


def sheet(name, widths):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    return ws


def put(ws, ref, value, font=F_N, fill=None, fmt=None, align=None, border=False):
    c = ws[ref]
    c.value = value
    c.font = font
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    if align:
        c.alignment = align
    if border:
        c.border = BOX
    return c


def header(ws, row, labels, start=1):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=start + i, value=lab)
        c.font, c.fill, c.border = F_HDR, FILL_HDR, BOX
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")


def name(n, ref):
    wb.defined_names[n] = DefinedName(n, attr_text=ref)


# =====================================================================
# README
# =====================================================================
ws = sheet("README", {"A": 2, "B": 22, "C": 110, "D": 2})
put(ws, "B2", "Chiller Energy Optimisation – Example Project (Excel MLR, REAL DATA)", F_TITLE)
put(ws, "B3", f"Client data for MT Chiller 1 ({P}) from '20261016_benny - Copy.xlsx': 5 periods of 2 days "
              f"(Feb 2024, Aug 2024, Feb 2025, Aug 2025, Feb 2026), 1-minute readings used exactly as recorded (no averaging) = {N:,} rows. "
              "Every result is an Excel formula: change the data or settings and everything recalculates.", F_N, align=WRAP)
ws.row_dimensions[3].height = 42
put(ws, "B5", "Workflow", F_H2)
flow = [("1. Raw_Data", "Raw 1-minute readings of the client's MT Chiller 1 tags (CL1, EF.PV, EF.PV2, FIT1-2, HR1, TIT1-4)"),
        ("2. Cleaning", "Calculate RT, chiller power (kW), heat balance → 5 checks → Keep / Remove + reason"),
        ("3. Clean_Data", "Only valid rows, in time order. First 80% (Hist1–Hist4) = Train, last 20% (Hist5, Feb 2026) = Test"),
        ("4. MLR_Model", "LINEST: Power (kW) = b0 + b1·RT + b2·TIT1 → coefficients, p-values, R², VIF, why other tags are left out"),
        ("5. Validation", "Test the model on unseen Feb 2026 data: R², CV(RMSE), NMBE (ASHRAE Guideline 14)"),
        ("6. Live_Data", "Stand-in live readings (1-minute values, 2 Feb 2026) → apply the trained constants → advice per snapshot"),
        ("7. Optimisation", "One live snapshot → try every allowed condenser water supply temperature (TIT1) → lowest kW = recommendation + saving"),
        ("8. Real_World_Guide", "What is still needed to go from this prototype to the delivered Python solution")]
r = 6
for k, v in flow:
    put(ws, f"B{r}", k, F_B, FILL_KEY)
    put(ws, f"C{r}", v)
    if r < 20:
        put(ws, f"B{r+1}", "        ↓")
    r += 2
put(ws, "B23", "Other sheets", F_H2)
put(ws, "B24", "Assumption", F_B); put(ws, "C24", "Every assumption behind the results, why it was made, and who must confirm it.")
put(ws, "B25", "Settings", F_B); put(ws, "C25", "All thresholds, limits and constants. Yellow cells are the ones you edit.")
put(ws, "B26", "Tag_List", F_B); put(ws, "C26", "The client's real tag names, what each one measures and its role in the model.")
put(ws, "B28", "Colour legend", F_H2)
put(ws, "B29", "Input (edit me)", F_IN, FILL_IN); put(ws, "C29", "Blue text on yellow: settings and live readings you can change")
put(ws, "B30", "Formula", F_N); put(ws, "C30", "Black text: calculated – do not overwrite")
put(ws, "B31", "Key result", F_B, FILL_KEY); put(ws, "C31", "Light blue: main outputs (coefficients, metrics)")
put(ws, "B32", "Recommendation", F_B, FILL_REC); put(ws, "C32", "Orange: the best option chosen by the optimisation")
put(ws, "B34", "What changed from the dummy version", F_H2)
changes = ["Now uses the raw 1-minute readings directly (previous version used 30-minute averages).", f"Dummy data (300 made-up snapshots) replaced by the client's real MT Chiller 1 data ({N:,} raw 1-minute readings).",
           "Tag roles corrected to the client's SCADA descriptions: TIT1/TIT2 = condenser water supply/return, TIT3/TIT4 = chilled water supply/return, FIT1 = condenser flow, FIT2 = chilled water flow.",
           "No chiller power tag is historised, so power (kW) = EF.PV2 × CL1 ÷ 3.517 (the SCADA efficiency script, sheet '3'). Cross-checked against EF.PV × 1,063 kW ÷ 100.",
           "Model changed from CL01 = b0 + b1·RT + b2·TIT01 + b3·TIT03 to Power = b0 + b1·RT + b2·TIT1. The chilled water supply (TIT3) is fixed at ~12.85 °C in all data, so it cannot be estimated or optimised.",
           "The lever is now the condenser water supply temperature TIT1 (set by the cooling towers), not the chilled water setpoint.",
           "The original dummy workbook is kept as 'ExampleProject (dummy version).xlsx'."]
for i, t in enumerate(changes):
    put(ws, f"B{35+i}", "•", F_B, align=Alignment(horizontal="right", vertical="top"))
    put(ws, f"C{35+i}", t, align=WRAP)
    ws.row_dimensions[35 + i].height = 28
put(ws, "B43", "Note: the client's final solution must be Python or PowerShell. This workbook is a prototype to prove the method on real data and explain it; each sheet maps to one step of the script.", F_I)

# =====================================================================
# Assumption
# =====================================================================
ws = sheet("Assumption", {"A": 7, "B": 58, "C": 62, "D": 22, "E": 40, "F": 22})
put(ws, "A1", "Project assumptions", F_TITLE)
put(ws, "A2", "What we assumed, why, and who must confirm it. Fill the yellow column as the client replies.", F_I)
put(ws, "A4", "Confirm these 5 first – they change the result the most", F_H2)
top5 = [("G1", "Lowest achievable condenser water supply temp (27 °C) – it decides every recommendation"),
        ("B3", "Chiller power = EF.PV2 × CL1 ÷ 3.517 (no power tag is historised)"),
        ("G9", "Extra cooling tower fan power when TIT1 is lowered (not in the model yet)"),
        ("I1", "Only 10 days of data – the full 3 years is needed for the final model"),
        ("H3", "Tariff 0.30 SGD/kWh and 8,760 h/yr")]
for i, (k, t) in enumerate(top5):
    put(ws, f"A{5+i}", k, F_B, FILL_REC); put(ws, f"B{5+i}", t, F_N, FILL_REC)
put(ws, "F4", "Status key", F_B)
for i, s in enumerate(["Confirmed (client file)", "Confirm with client", "Check with data", "Our decision", "Standard"]):
    put(ws, f"F{5+i}", s)
header(ws, 11, ["No.", "Assumption", "Why", "Status", "Client reply / comment"])
groups = [
    ("A. Tags & sensors", [
        ("A1", "TIT1 = condenser water SUPPLY (~29 °C), TIT2 = condenser water RETURN (~34 °C)", "SCADA description in sheet 2.Desc", "Confirmed (client file)"),
        ("A2", "TIT3 = chilled water SUPPLY (~12.85 °C), TIT4 = chilled water RETURN (~18 °C)", "SCADA description in sheet 2.Desc", "Confirmed (client file)"),
        ("A3", "FIT1 = condenser water flow, FIT2 = chilled water flow", "SCADA description; CL1 = FIT2 × (TIT4 − TIT3) × 1.163 matches within 0.2%", "Confirmed (client file)"),
        ("A4", "CL1.PV = cooling load (kW of heat removed), not electrical power", "Equals the chilled water energy balance; ~8,000 kW is far above the ~1,063 kW rated power", "Confirmed (client file)"),
        ("A5", "HR1.PV = heat rejected at the condenser (kW)", "Equals FIT1 × (TIT2 − TIT1) × 1.163 within 0.2%", "Check with data"),
        ("A6", "TIT2 is a RESULT of power, so it is not a predictor", "Heat rejected = cooling + power → TIT2 already contains the answer (leakage)", "Our decision"),
        ("A7", "Values are the raw 1-minute readings (no averaging)", "Uses every reading the client exported; more noise than 30-min averages, so R² is a little lower", "Our decision"),
    ]),
    ("B. Physics & constants", [
        ("B1", "Water constant = 1.163 kW per m³/h·°C", "4.186 kJ/kg·K × 1000 kg/m³ ÷ 3600 s", "Standard"),
        ("B2", "1 RT = 3.517 kW", "Definition of a refrigeration ton", "Standard"),
        ("B3", "Power (kW) = EF.PV2 × CL1 ÷ 3.517", "SCADA script: EF.PV2 = IMSValue ÷ (ChCLValue ÷ 3.517); IMSValue = power meter", "Confirm with client"),
        ("B4", "Rated power ≈ 1,063 kW", "EF.PV = power ÷ rated power × 100 → back-calculated, constant in all data", "Confirm with client"),
        ("B5", "Heat out ≈ cooling + power (±5%)", "Energy conservation; real data is within ~1%", "Check with data"),
        ("B6", "Reference efficiency = 0.4107 kW/RT", "Client reference value (tag_reference); average of the data is ~0.378", "Confirm with client"),
    ]),
    ("C. Data cleaning", [
        ("C1", "OFF when CL1 ≤ 500 kW or chilled water ΔT ≤ 0.5 °C", "No real cooling at these levels (Run status = 1 in all data)", "Check with data"),
        ("C2", "Valid TIT3 = 10–15 °C", "Chilled water setpoint ~12.85 °C; outside = sensor fault", "Check with data"),
        ("C3", "Valid TIT1 = 20–40 °C", "Normal condenser range in Singapore", "Check with data"),
        ("C4", "Plausible kW/RT = 0.30–0.60", "Real data 0.35–0.40; outside = meter fault or transient", "Check with data"),
        ("C5", "Each 1-minute reading is used as recorded", "No averaging, so no minimum-minutes rule is needed", "Our decision"),
        ("C6", "Remove a row if CL1, EF.PV2, TIT1, TIT3 or TIT4 is missing", "All needed for power, RT and the model", "Our decision"),
        ("C7", "Rows are independent of each other", "1-min data is strongly auto-correlated (each minute is close to the last), so p-values look smaller than they really are", "Check with data"),
    ]),
    ("D. Model choice", [
        ("D1", "Y = chiller power (kW)", "Client goal is energy saving, measured in kW", "Our decision"),
        ("D2", "X = RT (cooling load) and TIT1 (condenser water supply)", "Load is the demand; TIT1 is the condition the cooling towers control", "Our decision"),
        ("D3", "TIT3 not used", "Fixed setpoint: varies only 12.80–12.93 °C → effect cannot be estimated", "Check with data"),
        ("D4", "TIT4, FIT1, FIT2 not used", "TIT4 and FIT2 are already inside RT; FIT1 shows no stable link (sign flips between periods)", "Our decision"),
        ("D5", "EF.PV and EF.PV2 not used as X", "They are power itself (circular)", "Our decision"),
        ("D6", "Relationship is a straight line", "MLR assumption; checked by residuals in Validation", "Check with data"),
        ("D7", "No interaction terms (e.g. RT × TIT1)", "Simple and auditable; add if validation fails on 3-year data", "Check with data"),
        ("D8", "One model per chiller", "Each chiller has its own efficiency curve", "Our decision"),
    ]),
    ("E. Statistics", [
        ("E1", "Linearity", "Check: scatter of each X vs kW", "Check with data"),
        ("E2", "Errors independent", "Check: Durbin-Watson ≈ 2", "Check with data"),
        ("E3", "Errors normally distributed", "Check: residual histogram", "Check with data"),
        ("E4", "No multicollinearity", "Check: VIF < 5 (MLR_Model section 8)", "Check with data"),
        ("E5", "Significance level 5% (95% confidence)", "Common statistical standard", "Standard"),
    ]),
    ("F. Training & validation", [
        ("F1", "First 80% of time = Train (Hist1–Hist4), last 20% = Test (Hist5)", "Mimics real use: learn from the past, predict the future", "Our decision"),
        ("F2", "Pass = R² ≥ 0.75, CV(RMSE) ≤ 30%, |NMBE| ≤ 10%", "ASHRAE Guideline 14", "Standard"),
        ("F3", "Test R² is low-sensitive: Feb 2026 has a narrow load range", "R² falls when the test data varies little; CV(RMSE) is the better guide here", "Our decision"),
        ("F4", "Model valid only inside the training range", RANGE_TXT, "Our decision"),
    ]),
    ("G. Optimisation", [
        ("G1", "Lowest achievable TIT1 = 27.0 °C", f"Lowest seen {T1_LO:.1f} °C; real limit = wet-bulb + tower approach – ask the plant engineer", "Confirm with client"),
        ("G2", "Highest TIT1 in the search = 31.5 °C, step 0.5 °C", "Covers the observed range", "Our decision"),
        ("G3", "Load (RT) unchanged when TIT1 changes", "Factory demand does not depend on the condenser water", "Our decision"),
        ("G4", "Chilled water setpoint (TIT3) unchanged", "Not part of this optimisation", "Our decision"),
        ("G5", "Best TIT1 is always the lowest limit", "Model is linear in TIT1 (b2 > 0)", "Our decision"),
        ("G6", "Saving = model kW now − model kW at best", "Model vs model cancels model bias", "Our decision"),
        ("G7", "Saving ± = t × SE(b2) × |ΔTIT1|", "Only TIT1 changes, so only b2 uncertainty counts", "Standard"),
        ("G8", "No advice on invalid readings", "Never advise from bad data", "Our decision"),
        ("G9", "Cooling tower fan and pump power NOT included", "Lower TIT1 needs more fan power → net saving is smaller", "Confirm with client"),
        ("G10", "Advice only – operator decides", "Client requirement", "Standard"),
    ]),
    ("H. Reporting & cost", [
        ("H1", "kW/RT is chiller-only (not whole plant)", "No pump or tower fan power yet", "Confirm with client"),
        ("H2", "8,760 operating hours per year", "24/7 – upper-bound illustration", "Confirm with client"),
        ("H3", "Tariff = 0.30 SGD/kWh", "Placeholder", "Confirm with client"),
        ("H4", "Annual saving assumes same conditions all year", "Illustration, not a forecast", "Our decision"),
    ]),
    ("I. Data limits", [
        ("I1", "5 periods × 2 days = 10 days of data", "Enough to prove the method, not for the final model", "Confirm with client"),
        ("I2", "Only February and August of 2024–2026", "Other months (e.g. hottest May/June) not covered", "Check with data"),
        ("I3", "Live data = 16 one-minute readings on 2 Feb 2026", "Stand-in until a live feed is available", "Our decision"),
    ]),
]
r = 12
for g, rows in groups:
    put(ws, f"A{r}", g, F_B, FILL_SEC)
    for c in "BCDE":
        ws[f"{c}{r}"].fill = FILL_SEC
    r += 1
    for k, t, why, st in rows:
        put(ws, f"A{r}", k, F_B); put(ws, f"B{r}", t, align=WRAP); put(ws, f"C{r}", why, align=WRAP)
        put(ws, f"D{r}", st); put(ws, f"E{r}", None, F_IN, FILL_IN)
        r += 1
ws.freeze_panes = "A12"
for rng in [f"D12:D{r}"]:
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Confirm with client"'], font=Font(color="C00000", bold=True)))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Confirmed (client file)"'], font=Font(color="00703C", bold=True)))

# =====================================================================
# Settings
# =====================================================================
ws = sheet("Settings", {"A": 34, "B": 30, "C": 16, "D": 18, "E": 80})
put(ws, "A1", "Settings – constants, limits and thresholds", F_TITLE)
put(ws, "A2", "Edit only the yellow cells. Every other sheet reads these values by name.", F_I)
header(ws, 4, ["Parameter", "Value", "Unit", "Name used in formulas", "Source / note"])
settings = [
    ("Chiller ID", P, "", None, "MT Chiller 1. One workbook (or one model) per chiller."),
    ("Rated power", 1063, "kW", "Rated_kW", "Back-calculated from EF.PV = power ÷ rated power × 100. Confirm on the datasheet."),
    ("Reference efficiency", 0.4107, "kW/RT", "Ref_kWRT", "Client reference value for EF.PV2."),
    ("Water heat constant", 1.163, "kW per m³/h·°C", "Water_K", "Physics: 4.186 kJ/kg·K × 1000 kg/m³ ÷ 3600 s."),
    ("1 RT in kW", 3.517, "kW", "RT_kW", "Definition of a refrigeration ton."),
    ("Minimum cooling load", 500, "kW", "MinCL", "Below this the chiller is treated as OFF."),
    ("Minimum chilled water ΔT", 0.5, "°C", "MinDT", "Below this there is no real cooling."),
    ("TIT3 valid low", 10, "°C", "CHWS_Lo", "Sensor sanity range for chilled water supply."),
    ("TIT3 valid high", 15, "°C", "CHWS_Hi", None),
    ("TIT1 valid low", 20, "°C", "CWS_Lo", "Sensor sanity range for condenser water supply."),
    ("TIT1 valid high", 40, "°C", "CWS_Hi", None),
    ("kW/RT plausible low", 0.30, "kW/RT", "KWRT_Lo", "Outside this band = outlier."),
    ("kW/RT plausible high", 0.60, "kW/RT", "KWRT_Hi", None),
    ("Heat balance tolerance", 0.05, "fraction", "HB_Tol", "|HR1 − (cooling + power)| ÷ (cooling + power)."),
    ("Training share", 0.8, "fraction", "TrainFrac", "First 80% of clean rows (in time order) train the model, last 20% test it."),
    ("Setpoint search – lowest TIT1", 27.0, "°C", "SP_Min", "Assumption – lowest condenser water the towers can reach. Confirm with the plant engineer."),
    ("Setpoint search – highest TIT1", 31.5, "°C", "SP_Max", "Covers the observed range."),
    ("Setpoint search – step", 0.5, "°C", "SP_Step", None),
    ("Confidence level", 0.95, "fraction", "Conf", "For confidence and prediction intervals."),
    ("Operating hours per year", 8760, "h", "Hours_Yr", "24 × 365 – used for the annual illustration only."),
    ("Electricity tariff", 0.30, "SGD/kWh", "Tariff", "Placeholder – replace with the client's actual tariff."),
]
for i, (p, v, u, n, note) in enumerate(settings):
    r = 5 + i
    put(ws, f"A{r}", p); put(ws, f"B{r}", v, F_IN, FILL_IN, align=Alignment(horizontal="left"))
    put(ws, f"C{r}", u); put(ws, f"D{r}", n); put(ws, f"E{r}", note, F_I)
    if n:
        name(n, f"Settings!$B${r}")

# =====================================================================
# Tag_List
# =====================================================================
ws = sheet("Tag_List", {"A": 40, "B": 30, "C": 46, "D": 9, "E": 16, "F": 46})
put(ws, "A1", "Tag list – what each sensor measures (MT Chiller 1)", F_TITLE)
header(ws, 3, ["Tag", "Simple name", "What it measures (client SCADA description)", "Unit", "Typical value", "Role in this project"])
tags = [
    (P + "_CL1.PV", "Cooling load", "Heat removed from the chilled water", "kW", "~8,000", "Converted to RT → X1 (the demand to meet)"),
    (P + "_EF.PV", "Power loading", "Power ÷ rated power × 100", "%", "~80", "Cross-check of the power calculation"),
    (P + "_EF.PV2", "Efficiency", "Power ÷ cooling (RT)", "kW/RT", "~0.38", "Used to calculate power (Y); reported KPI"),
    (P + "_FIT1.PV", "Condenser water flow", "Condenser water return flow meter", "m³/h", "~1,600", "Not used (no stable link)"),
    (P + "_FIT2.PV", "Chilled water flow", "Chilled water supply flow meter", "m³/h", "~1,250", "Inside CL1 (data check only)"),
    (P + "_HR1.PV", "Heat rejection", "Heat rejected at the condenser", "kW", "~8,800", "Heat balance check (data quality)"),
    (P + "_TIT1.PV", "Condenser water supply temp", "Water from the cooling towers into the chiller", "°C", "~29", "X2 – the lever we optimise (via cooling towers)"),
    (P + "_TIT2.PV", "Condenser water return temp", "Water from the chiller back to the towers", "°C", "~34", "Not used (result of power – leakage)"),
    (P + "_TIT3.PV", "Chilled water supply temp", "Chilled water leaving the chiller", "°C", "~12.85", "Validity check only (fixed setpoint)"),
    (P + "_TIT4.PV", "Chilled water return temp", "Chilled water coming back from the plant", "°C", "~18", "Inside CL1 (data check only)"),
    ("RT (calculated)", "Cooling load in tons", "CL1 ÷ 3.517", "RT", "~2,280", "X1 – the demand to meet"),
    ("Power (calculated)", "Chiller electrical power", "EF.PV2 × RT", "kW", "~860", "Y – the output the model predicts"),
]
for i, row in enumerate(tags):
    for j, v in enumerate(row):
        put(ws, f"{'ABCDEF'[j]}{4+i}", v, F_B if j == 0 else F_N, border=True, align=WRAP)
put(ws, "A17", "Source: sheets 2.Desc and 3 of '20261016_benny - Copy.xlsx'. Power is not historised directly; it comes from the SCADA efficiency script.", F_I)

# =====================================================================
# Raw_Data
# =====================================================================
ws = sheet("Raw_Data", {"A": 18, "B": 9, **{c: 11 for c in "CDEFGHIJKL"}, "M": 9, "N": 3, "O": 90})
hdrs = ["Timestamp (1-min)", "Period", "CL1.PV (kW)", "EF.PV (%)", "EF.PV2 (kW/RT)", "FIT1.PV (m³/h)", "FIT2.PV (m³/h)",
        "HR1.PV (kW)", "TIT1.PV (°C)", "TIT2.PV (°C)", "TIT3.PV (°C)", "TIT4.PV (°C)"]
header(ws, 1, hdrs)
put(ws, "O1", "About this sheet", F_HDR, FILL_HDR)
fmts = ["yyyy-mm-dd hh:mm", "@", "#,##0.0", "0.00", "0.0000", "#,##0.0", "#,##0.0", "#,##0.0", "0.00", "0.00", "0.000", "0.000"]
for i, rec in a.iterrows():
    r = i + 2
    vals = [rec.slot.to_pydatetime(), rec.Period] + [None if pd.isna(rec[c]) else round(float(rec[c]), 5) for c in cols]
    for j, v in enumerate(vals):
        c = ws.cell(row=r, column=j + 1, value=v)
        c.font, c.number_format = F_N, fmts[j]
about = [f"Client data, MT Chiller 1: sheets '2.Hist Value1' to '2.Hist Value5' of 20261016_benny - Copy.xlsx.",
         "Each row = one raw 1-minute reading, exactly as exported by the client (no averaging).",
         "Hist1 01–02 Feb 2024 | Hist2 01–02 Aug 2024 | Hist3 01–02 Feb 2025 | Hist4 01–02 Aug 2025 | Hist5 01–02 Feb 2026.",
         f"{N:,} rows. There is no chiller power tag – power is calculated on the Cleaning sheet from EF.PV2 and CL1.",
         "To use the full 3-year data: paste it here in the same column order and extend the formulas (Real_World_Guide)."]
for i, t in enumerate(about):
    put(ws, f"O{2+i}", t, F_I)
ws.freeze_panes = "C2"

# =====================================================================
# Cleaning
# =====================================================================
ws = sheet("Cleaning", {"A": 17, **{c: 11 for c in "BCDEFG"}, **{c: 10 for c in "HIJKL"}, "M": 8, "N": 16, "O": 8, "P": 9, "Q": 44, "R": 9, "S": 9})
header(ws, 1, ["Timestamp", "CHW ΔT (°C)", "RT", "Power (kW)", "Heat in (kW)", "Heat out HR1 (kW)", "Heat balance error",
               "1. Complete?", "2. Running?", "3. In range?", "4. Heat balance OK?", "5. kW/RT plausible?", "KEEP?", "Reason", "Keep No.", "Kept so far"])
R = "Raw_Data!"
for r in range(2, LAST + 1):
    f = {
        "A": f"={R}A{r}",
        "B": f'=IF(OR({R}K{r}="",{R}L{r}=""),"",{R}L{r}-{R}K{r})',
        "C": f'=IF({R}C{r}="","",{R}C{r}/RT_kW)',
        "D": f'=IF(OR(C{r}="",{R}E{r}=""),"",{R}E{r}*C{r})',
        "E": f'=IF(OR({R}C{r}="",D{r}=""),"",{R}C{r}+D{r})',
        "F": f'=IF({R}H{r}="","",{R}H{r})',
        "G": f'=IF(OR(E{r}="",F{r}=""),"",IF(E{r}<=0,"",(F{r}-E{r})/E{r}))',
        "H": f"=COUNT({R}C{r},{R}E{r},{R}I{r},{R}K{r},{R}L{r})=5",
        "I": f"=IF(H{r},AND({R}C{r}>MinCL,B{r}>MinDT),FALSE)",
        "J": f"=IF(H{r},AND({R}K{r}>=CHWS_Lo,{R}K{r}<=CHWS_Hi,{R}I{r}>=CWS_Lo,{R}I{r}<=CWS_Hi),FALSE)",
        "K": f'=IF(G{r}="",TRUE,ABS(G{r})<=HB_Tol)',
        "L": f"=IF(AND(H{r},I{r},J{r}),AND({R}E{r}>=KWRT_Lo,{R}E{r}<=KWRT_Hi),FALSE)",
        "M": f"=AND(H{r},I{r},J{r},K{r},L{r})",
        "N": f'=IF(M{r},"OK",IF(NOT(H{r}),"Missing value",IF(NOT(I{r}),"Chiller off",IF(NOT(J{r}),"Out of range",IF(NOT(K{r}),"Heat balance fail","kW/RT outlier")))))',
        "O": f'=IF(M{r},P{r},"")',
        "P": f"=N(P{r-1})+IF(M{r},1,0)",
    }
    for col, v in f.items():
        c = ws[f"{col}{r}"]; c.value = v; c.font = F_N
    ws[f"A{r}"].number_format = "yyyy-mm-dd hh:mm"
    for col, fm in zip("BCDEFG", ["0.00", "#,##0.0", "#,##0.0", "#,##0.0", "#,##0.0", "0.0%"]):
        ws[f"{col}{r}"].number_format = fm
ws.conditional_formatting.add(f"M2:M{LAST}", CellIsRule(operator="equal", formula=["FALSE"], fill=PatternFill("solid", fgColor="F8CBAD")))
header(ws, 2, ["Cleaning summary", "Rows", "Share"], start=17)
summ = [("Total raw rows", f"=COUNTA({R}A2:A{LAST})"), ("Kept (clean)", f"=COUNTIF(M2:M{LAST},TRUE)"), ("Removed", "=R3-R4")] + \
       [(f"  – {t}", f'=COUNTIF(N2:N{LAST},"{t}")') for t in ["Missing value", "Chiller off", "Out of range", "Heat balance fail", "kW/RT outlier"]]
for i, (t, fx) in enumerate(summ):
    r = 3 + i
    put(ws, f"Q{r}", t, F_B if i < 3 else F_N, FILL_KEY if i == 1 else None)
    put(ws, f"R{r}", fx, F_B if i == 1 else F_N, FILL_KEY if i == 1 else None)
    put(ws, f"S{r}", f"=R{r}/$R$3", fmt="0.0%")
name("N_Keep", "Cleaning!$R$4")
notes = ["Checks run in order; the first failed check is the Reason.",
         "1 Complete: CL1, EF.PV2, TIT1, TIT3, TIT4 all present in that minute",
         "2 Running: cooling load > MinCL and chilled water ΔT > MinDT",
         "3 In range: TIT3 within CHWS_Lo–Hi and TIT1 within CWS_Lo–Hi",
         "4 Heat balance: |HR1 − (cooling + power)| ÷ (cooling + power) ≤ HB_Tol",
         "5 kW/RT (EF.PV2) within KWRT_Lo–Hi",
         "Power (kW) = EF.PV2 × RT – the SCADA efficiency script run backwards.",
         "Real data note: the client data is already clean (chiller ran the whole time), so almost nothing is removed. The checks matter for the 3-year data."]
for i, t in enumerate(notes):
    put(ws, f"Q{12+i}", t, F_I if i else F_B)
ws.freeze_panes = "B2"

# =====================================================================
# Clean_Data
# =====================================================================
ws = sheet("Clean_Data", {"A": 6, "B": 8, "C": 17, "D": 8, "E": 12, "F": 11, "G": 11, "H": 10, "I": 10, "J": 11, "K": 11,
                          "L": 9, "M": 8, "N": 12, "O": 11, "P": 11, "Q": 3, "R": 90})
header(ws, 1, ["No.", "Source row", "Timestamp", "Period", "Power (kW)  [Y]", "RT  [X1]", "TIT1 (°C)  [X2]", "TIT2 (°C)",
               "TIT3 (°C)", "TIT4 (°C)", "Actual kW/RT", "EF.PV (%)", "Set", "Predicted kW", "Residual (kW)", "Predicted kW/RT"])
put(ws, "R1", "About this sheet", F_HDR, FILL_HDR)
C = "Cleaning!"
rng = lambda col: f"${col}$2:${col}${LAST}"
for r in range(2, LAST + 1):
    f = {
        "A": f'=IF(ROW()-1<=N_Keep,ROW()-1,"")',
        "B": f'=IF($A{r}="","",IFERROR(MATCH(A{r}-0.5,{C}$P$2:$P${LAST},1),0)+1)',
        "C": f'=IF($A{r}="","",INDEX({C}{rng("A")},$B{r}))',
        "D": f'=IF($A{r}="","",INDEX({R}{rng("B")},$B{r}))',
        "E": f'=IF($A{r}="","",INDEX({C}{rng("D")},$B{r}))',
        "F": f'=IF($A{r}="","",INDEX({C}{rng("C")},$B{r}))',
        "G": f'=IF($A{r}="","",INDEX({R}{rng("I")},$B{r}))',
        "H": f'=IF($A{r}="","",INDEX({R}{rng("J")},$B{r}))',
        "I": f'=IF($A{r}="","",INDEX({R}{rng("K")},$B{r}))',
        "J": f'=IF($A{r}="","",INDEX({R}{rng("L")},$B{r}))',
        "K": f'=IF($A{r}="","",INDEX({R}{rng("E")},$B{r}))',
        "L": f'=IF($A{r}="","",INDEX({R}{rng("D")},$B{r}))',
        "M": f'=IF($A{r}="","",IF(A{r}<=N_Train,"Train","Test"))',
        "N": f'=IF($A{r}="","",Coef_Int+Coef_RT*F{r}+Coef_TIT1*G{r})',
        "O": f'=IF($A{r}="","",E{r}-N{r})',
        "P": f'=IF($A{r}="","",N{r}/F{r})',
    }
    for col, v in f.items():
        c = ws[f"{col}{r}"]; c.value = v; c.font = F_N
    for col, fm in zip("CEFGHIJKLNOP", ["yyyy-mm-dd hh:mm", "#,##0.0", "#,##0.0", "0.00", "0.00", "0.000", "0.000", "0.0000", "0.00", "#,##0.0", "#,##0.0", "0.0000"]):
        ws[f"{col}{r}"].number_format = fm
ws.conditional_formatting.add(f"M2:M{LAST}", CellIsRule(operator="equal", formula=['"Test"'], fill=PatternFill("solid", fgColor="FCE4D6")))
for i, t in enumerate(["Only rows with KEEP = TRUE, pulled in time order with INDEX/MATCH on Cleaning!P (running count of kept rows).",
                       "Columns E–G are the regression data: Y = Power, X = RT, TIT1 (X must be side by side for LINEST).",
                       "TIT2, TIT3, TIT4 are kept only to show why they are NOT used (MLR_Model section 7).",
                       "Set: first N_Train rows = Train (fit the model), the rest = Test (check it on unseen data).",
                       "With all rows kept: Train = Hist1–Hist4 (11,520 minutes), Test = Hist5, Feb 2026 (2,880 minutes)."]):
    put(ws, f"R{2+i}", t, F_I)
ws.freeze_panes = "D2"
name("Y_Train", "OFFSET(Clean_Data!$E$2,0,0,N_Train,1)")
name("X_Train", "OFFSET(Clean_Data!$F$2,0,0,N_Train,2)")
name("X_RT", "OFFSET(Clean_Data!$F$2,0,0,N_Train,1)")
name("X_T1", "OFFSET(Clean_Data!$G$2,0,0,N_Train,1)")
name("C_T2", "OFFSET(Clean_Data!$H$2,0,0,N_Train,1)")
name("C_T3", "OFFSET(Clean_Data!$I$2,0,0,N_Train,1)")
name("C_T4", "OFFSET(Clean_Data!$J$2,0,0,N_Train,1)")
name("Y_Test", "OFFSET(Clean_Data!$E$2,N_Train,0,N_Test,1)")
name("P_Test", "OFFSET(Clean_Data!$N$2,N_Train,0,N_Test,1)")

# =====================================================================
# MLR_Model
# =====================================================================
ws = sheet("MLR_Model", {"A": 32, "B": 15, "C": 15, "D": 15, "E": 15, "F": 13, "G": 34, "H": 62})
put(ws, "A1", "Multiple Linear Regression – trained with LINEST on real data", F_TITLE)
put(ws, "A3", "1. Model", F_H2)
put(ws, "A4", "General form", F_B); put(ws, "B4", "Y = b0 + b1·X1 + b2·X2 + ε")
put(ws, "A5", "This project", F_B); put(ws, "B5", "Power (kW) = b0 + b1·RT + b2·TIT1")
put(ws, "A6", "Why", F_B)
put(ws, "B6", "Y (power) is continuous; the 2 inputs are continuous; goal = predict kW so condenser water temperatures can be compared. "
              "RT is the demand, TIT1 is the condenser water from the cooling towers (the lever).", align=WRAP)
ws.merge_cells("B6:H6"); ws.row_dimensions[6].height = 30
put(ws, "A8", "2. Training data", F_H2)
for i, (t, fx, fm) in enumerate([("Clean rows available", "=N_Keep", "0"), ("Training rows (first share)", "=ROUND(N_Keep*TrainFrac,0)", "0"),
                                  ("Test rows (the rest)", "=N_Keep-B10", "0"),
                                  ("Training period from", f"=INDEX(Clean_Data!$C$2:$C${LAST},1)", "yyyy-mm-dd hh:mm"),
                                  ("Training period to", f"=INDEX(Clean_Data!$C$2:$C${LAST},N_Train)", "yyyy-mm-dd hh:mm")]):
    put(ws, f"A{9+i}", t); put(ws, f"B{9+i}", fx, fmt=fm, align=Alignment(horizontal="left"))
name("N_Train", "MLR_Model!$B$10"); name("N_Test", "MLR_Model!$B$11")
put(ws, "A15", "3. LINEST raw output   =LINEST(Y_Train, X_Train, TRUE, TRUE)", F_H2)
put(ws, "A16", "LINEST returns the X coefficients in REVERSE order (last X first). Blank cells are #N/A in Excel by design.", F_I)
header(ws, 17, ["Row", "TIT1 (b2)", "RT (b1)", "Intercept (b0)"])
L = "LINEST(Y_Train,X_Train,TRUE,TRUE)"
for i, t in enumerate(["Coefficient", "Standard error", "R²  |  SE of Y", "F  |  df (residual)", "SS regression  |  SS residual"]):
    r = 18 + i
    put(ws, f"A{r}", t)
    ncol = 3 if i < 2 else 2
    for j in range(ncol):
        put(ws, f"{'BCD'[j]}{r}", f"=INDEX({L},{i+1},{j+1})", fmt="0.0000")
put(ws, "A24", "4. Coefficients", F_H2)
header(ws, 25, ["Term", "Coefficient", "Std error", "t stat", "p-value", "Significant?", "Effect", "Meaning"])
terms = [("Intercept (b0)", "D", "—", "Value when all X = 0. No physical meaning on its own."),
         ("RT (b1)", "C", '="+100 RT load → "&TEXT(B27*100,"+0.0;-0.0")&" kW"', "Higher load needs more power (expected: positive)."),
         ("TIT1 (b2)", "B", '="+1 °C condenser water → "&TEXT(B28,"+0.0;-0.0")&" kW"', "Warmer condenser water = more compressor work (expected: positive). This is the lever we optimise.")]
for i, (t, col, eff, mean) in enumerate(terms):
    r = 26 + i
    put(ws, f"A{r}", t, F_B)
    put(ws, f"B{r}", f"={col}18", F_B, FILL_KEY, "0.0000")
    put(ws, f"C{r}", f"={col}19", fmt="0.0000")
    put(ws, f"D{r}", f"=B{r}/C{r}", fmt="0.00")
    put(ws, f"E{r}", f"=TDIST(ABS(D{r}),DF_Res,2)", fmt="0.0E+00")
    put(ws, f"F{r}", f'=IF(E{r}<1-Conf,"Yes","No")')
    put(ws, f"G{r}", eff); put(ws, f"H{r}", mean, align=WRAP)
name("Coef_Int", "MLR_Model!$B$26"); name("Coef_RT", "MLR_Model!$B$27"); name("Coef_TIT1", "MLR_Model!$B$28"); name("SE_TIT1", "MLR_Model!$C$28")
put(ws, "A30", "5. Model fit (training data)", F_H2)
fit = [("R²", "=B20", "0.0000", "Share of the variation in kW explained by the model (1 = perfect)."),
       ("Adjusted R²", "=1-(1-B31)*(N_Train-1)/(N_Train-2-1)", "0.0000", "R² corrected for the number of inputs."),
       ("Standard error of Y (SEy)", "=C20", "0.00", "Typical prediction error in kW."),
       ("F statistic", "=B21", "0.0", "Tests whether the model as a whole is useful."),
       ("Residual degrees of freedom", "=C21", "0", "n − k − 1"),
       ("p-value of F", "=FDIST(B34,2,B35)", "0.0E+00", "< 0.05 → the model is significant."),
       ("t critical (two-sided)", "=TINV(1-Conf,DF_Res)", "0.000", "Used for the ± intervals.")]
for i, (t, fx, fm, m) in enumerate(fit):
    r = 31 + i
    put(ws, f"A{r}", t); put(ws, f"B{r}", fx, F_B if i in (0, 2) else F_N, FILL_KEY if i in (0, 2) else None, fm); put(ws, f"C{r}", m, F_I)
name("SEy", "MLR_Model!$B$33"); name("DF_Res", "MLR_Model!$B$35"); name("T_Crit", "MLR_Model!$B$37")
put(ws, "A39", "6. Fitted equation (the constants you apply to live data)", F_H2)
put(ws, "A40", '="Power (kW) = "&TEXT(Coef_Int,"0.0")&" "&IF(Coef_RT<0,"− ","+ ")&TEXT(ABS(Coef_RT),"0.0000")&"·RT "&IF(Coef_TIT1<0,"− ","+ ")&TEXT(ABS(Coef_TIT1),"0.00")&"·TIT1"',
    Font(name="Arial", size=13, bold=True, color=NAVY), FILL_REC)
ws.merge_cells("A40:H40")
put(ws, "A41", '="In words: each +1 °C of condenser water supply costs about "&TEXT(Coef_TIT1,"0")&" kW ("&TEXT(Coef_TIT1/AVERAGE(Y_Train),"0.0%")&" of average power) at the same load."', F_I)
put(ws, "A43", "7. Correlation matrix (training data) – which tags go into the model and why", F_H2)
labels = ["Power", "RT", "TIT1", "TIT2", "TIT3", "TIT4"]
rngs = ["Y_Train", "X_RT", "X_T1", "C_T2", "C_T3", "C_T4"]
header(ws, 44, [""] + labels)
for i, (li, ri) in enumerate(zip(labels, rngs)):
    r = 45 + i
    put(ws, f"A{r}", li, F_B)
    for j, rj in enumerate(rngs):
        put(ws, f"{'BCDEFG'[j]}{r}", f"=CORREL({ri},{rj})", fmt="0.000", border=True)
ws.conditional_formatting.add("B45:G50", CellIsRule(operator="greaterThanOrEqual", formula=["0.8"], fill=PatternFill("solid", fgColor="9DC3E6")))
ws.conditional_formatting.add("B45:G50", CellIsRule(operator="between", formula=["0.5", "0.8"], fill=PatternFill("solid", fgColor="DDEBF7")))
why = [("TIT1", "IN – condenser water supply, set by the cooling towers. Strong link to power and only moderately linked to RT."),
       ("RT", "IN – the cooling demand; power must rise with load."),
       ("TIT2", "OUT – condenser water return = TIT1 + heat rejected, and heat rejected includes the power (leakage). r with TIT1 ≈ 0.97."),
       ("TIT3", '=IF(STDEV(C_T3)<0.1,"OUT – fixed chilled water setpoint: std dev only "&TEXT(STDEV(C_T3),"0.000")&" °C, so its effect cannot be estimated.","Check – TIT3 now varies, consider adding it.")'),
       ("TIT4", "OUT – already inside RT (cooling load = flow × (TIT4 − TIT3))."),
       ("FIT1 / FIT2", "OUT – FIT2 is inside RT; FIT1 changes sign between periods (CL_Correlation_Hist1-5.xlsx).")]
put(ws, "A52", "Decision", F_B)
for i, (t, w) in enumerate(why):
    put(ws, f"A{53+i}", t, F_B); put(ws, f"B{53+i}", w)
put(ws, "A60", "8. Multicollinearity check", F_H2)
header(ws, 61, ["Input", "R² vs other X", "VIF", "Verdict"])
for i, t in enumerate(["RT", "TIT1"]):
    r = 62 + i
    put(ws, f"A{r}", t, F_B); put(ws, f"B{r}", "=D46^2", fmt="0.000"); put(ws, f"C{r}", f"=1/(1-B{r})", F_B, FILL_KEY, "0.00")
    put(ws, f"D{r}", f'=IF(C{r}<5,"OK (< 5)",IF(C{r}<10,"Watch (5–10)","Problem (> 10)"))')
put(ws, "A64", "With 2 inputs, VIF = 1 ÷ (1 − r²) where r = correlation between RT and TIT1. Above 5: the coefficients become unstable.", F_I)

# =====================================================================
# Validation
# =====================================================================
TMAX = 3000  # room for up to 3,000 test rows (Hist5 = 2,880 minutes)
VL = 4 + TMAX
ws = sheet("Validation", {"A": 34, "B": 13, "C": 9, "D": 9, "E": 52, "F": 3, "G": 6, "H": 17, "I": 11, "J": 11, "K": 10, "L": 10,
                          "M": 9, "N": 10, "O": 3, "P": 10, "Q": 10, "R": 3, "S": 10, "T": 10, "U": 10, "V": 10})
put(ws, "A1", "Validation – does the model work on data it has not seen?", F_TITLE)
put(ws, "A2", "Test set = the last 20% of clean rows (Hist5, 1–2 Feb 2026, 2,880 minutes). Criteria from ASHRAE Guideline 14 (written for hourly data – applying them to 1-minute data is a stricter test).", F_I)
put(ws, "G1", "Actual vs predicted – every test snapshot", F_H2)
put(ws, "G2", "Error = Actual − Predicted.   Orange = model predicted too LOW.   Blue = model predicted too HIGH.", F_I)
header(ws, 4, ["Metric", "Value", "Target", "Pass?", "Meaning"])
header(ws, 4, ["No.", "Timestamp", "Actual kW", "Predicted kW", "Error kW", "|Error| kW", "Error %", "Tolerance"], start=7)
header(ws, 4, ["Chart: too low", "Chart: too high"], start=16)
header(ws, 4, ["Actual", "Perfect", "+ tol", "− tol"], start=19)
put(ws, "P3", "Chart helpers", F_I); put(ws, "S3", "Chart helpers: diagonal", F_I)
met = [("Test rows", "=N_Test", None, None, "Number of unseen 1-minute rows.", "0"),
       ("Mean actual kW (test)", "=AVERAGE(Y_Test)", None, None, None, "#,##0.0"),
       ("Sum of squared errors", "=SUMXMY2(Y_Test,P_Test)", None, None, None, "#,##0"),
       ("R² (test)", "=1-B7/DEVSQ(Y_Test)", 0.75, '=IF(B8>=C8,"PASS","FAIL")', "Share of variation explained on new data. Low when the test period barely varies.", "0.000"),
       ("CV(RMSE)", "=SQRT(B7/(N_Test-3))/B6", 0.3, '=IF(B9<=C9,"PASS","FAIL")', "Typical error as % of mean kW. ≤ 30% hourly (aim ≤ 15%).", "0.0%"),
       ("NMBE", "=(SUM(Y_Test)-SUM(P_Test))/((N_Test-3)*B6)", 0.1, '=IF(ABS(B10)<=C10,"PASS","FAIL")', "Bias: does it over/under-predict on average? within ±10%.", "0.0%"),
       ("R² (training, for comparison)", "=MLR_Model!B31", None, None, "Training data covers a wider range, so its R² is higher.", "0.000")]
for i, (t, fx, tgt, ps, m, fm) in enumerate(met):
    r = 5 + i
    put(ws, f"A{r}", t); put(ws, f"B{r}", fx, F_B if tgt else F_N, FILL_KEY if tgt else None, fm)
    if tgt is not None:
        put(ws, f"C{r}", tgt, F_IN, FILL_IN, fm); put(ws, f"D{r}", ps, F_B)
    put(ws, f"E{r}", m, F_I)
ws.conditional_formatting.add("D8:D10", CellIsRule(operator="equal", formula=['"PASS"'], font=Font(color="00703C", bold=True)))
ws.conditional_formatting.add("D8:D10", CellIsRule(operator="equal", formula=['"FAIL"'], font=Font(color="C00000", bold=True)))
put(ws, "A13", "Overall", F_B); put(ws, "B13", '=IF(COUNTIF(D8:D10,"PASS")=3,"MODEL ACCEPTED","REVIEW MODEL")', F_B, FILL_REC)
put(ws, "A15", "Power calculation cross-check (all clean rows)", F_H2)
header(ws, 16, ["Check", "Value", "", "", "Conclusion"])
xc = [("Rows checked", f"=N_Keep", "0", None),
      ("Mean of power ÷ (EF.PV × Rated_kW ÷ 100)", f"=SUMPRODUCT(Clean_Data!E2:E{LAST}/(Clean_Data!L2:L{LAST}*Rated_kW/100)*(Clean_Data!A2:A{LAST}<>\"\"))/N_Keep", "0.0000",
       '=IF(ABS(B18-1)<0.01,"≈ 1.000 → EF.PV2 × RT and EF.PV × rated power give the same kW: the power calculation is consistent","Mismatch – check rated power")'),
      ("Mean heat balance error", f"=AVERAGE(Cleaning!G2:G{LAST})", "0.0%", '="HR1 is "&TEXT(-B19,"0.0%")&" below cooling + power on average – within tolerance, sensors agree"'),
      ("Average kW/RT (clean data)", f"=AVERAGE(Clean_Data!K2:K{LAST})", "0.0000", '="vs reference "&TEXT(Ref_kWRT,"0.0000")&" kW/RT → chiller runs "&TEXT(1-B20/Ref_kWRT,"0.0%")&" better than the reference"')]
for i, (t, fx, fm, concl) in enumerate(xc):
    r = 17 + i
    put(ws, f"A{r}", t); put(ws, f"B{r}", fx, fmt=fm); put(ws, f"E{r}", concl)
from openpyxl.worksheet.formula import ArrayFormula
ws["B18"] = ArrayFormula("B18", f'=AVERAGE(IF(Clean_Data!A2:A{LAST}<>"",Clean_Data!E2:E{LAST}/(Clean_Data!L2:L{LAST}*Rated_kW/100)))')
put(ws, "A22", "Prediction error summary (test set)", F_H2)
header(ws, 23, ["Measure", "Value", "", "", "Meaning"])
put(ws, "A24", "Tolerance band"); put(ws, "B24", 0.02, F_IN, FILL_IN, "0.0%"); put(ws, "E24", "Edit me: how close counts as a good prediction (used for colours and the band).", F_I)
name("Err_Tol", "Validation!$B$24")
for i, (t, fx, fm, m) in enumerate([("Mean absolute error (kW)", f"=AVERAGE(L5:L{VL})", "0.0", "Average size of the miss, ignoring direction."),
                                     ("Largest miss (kW)", f"=MAX(L5:L{VL})", "0.0", "Worst single minute."),
                                     ("Mean error / bias (kW)", f"=AVERAGE(K5:K{VL})", "0.0", "Near 0 = no systematic over- or under-prediction."),
                                     ("Rows within tolerance", f'=COUNTIF(N5:N{VL},"Within")/N_Test', "0.0%", "Share of predictions inside the ± band.")]):
    r = 25 + i
    put(ws, f"A{r}", t); put(ws, f"B{r}", fx, F_B, FILL_KEY, fm); put(ws, f"E{r}", m, F_I)
for k in range(1, TMAX + 1):
    r = 4 + k
    f = {"G": f'=IF({k}<=N_Test,{k},"")',
         "H": f'=IF($G{r}="","",INDEX(Clean_Data!$C$2:$C${LAST},N_Train+G{r}))',
         "I": f'=IF($G{r}="","",INDEX(Clean_Data!$E$2:$E${LAST},N_Train+G{r}))',
         "J": f'=IF($G{r}="","",INDEX(Clean_Data!$N$2:$N${LAST},N_Train+G{r}))',
         "K": f'=IF($G{r}="","",I{r}-J{r})', "L": f'=IF($G{r}="","",ABS(K{r}))', "M": f'=IF($G{r}="","",K{r}/I{r})',
         "N": f'=IF($G{r}="","",IF(ABS(M{r})<=Err_Tol,"Within","Outside"))',
         "P": f'=IF($G{r}="",NA(),IF(K{r}>0,K{r},0))', "Q": f'=IF($G{r}="",NA(),IF(K{r}<0,K{r},0))'}
    for col, v in f.items():
        c = ws[f"{col}{r}"]; c.value = v; c.font = F_N
    for col, fm in zip("HIJKLM", ["dd-mmm hh:mm", "#,##0.0", "#,##0.0", "+0.0;-0.0", "0.0", "+0.0%;-0.0%"]):
        ws[f"{col}{r}"].number_format = fm
ws.conditional_formatting.add(f"K5:K{VL}", CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor="F8CBAD")))
ws.conditional_formatting.add(f"K5:K{VL}", CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor="BDD7EE")))
ws.conditional_formatting.add(f"N5:N{VL}", CellIsRule(operator="equal", formula=['"Outside"'], font=Font(color="C00000", bold=True)))
put(ws, "S5", f"=MIN(I5:I{VL})", fmt="0"); put(ws, "T5", "=S5", fmt="0"); put(ws, "U5", "=S5*(1+Err_Tol)", fmt="0"); put(ws, "V5", "=S5*(1-Err_Tol)", fmt="0")
put(ws, "S6", f"=MAX(I5:I{VL})", fmt="0"); put(ws, "T6", "=S6", fmt="0"); put(ws, "U6", "=S6*(1+Err_Tol)", fmt="0"); put(ws, "V6", "=S6*(1-Err_Tol)", fmt="0")
ws.freeze_panes = "A5"
NT = 2880  # rows shown in charts (Hist5, every minute)
ch = ScatterChart(); ch.title = "Actual vs predicted power (test, Feb 2026)"; ch.style = 13
ch.x_axis.title = "Actual kW"; ch.y_axis.title = "Predicted kW"
s = Series(Reference(ws, min_col=10, min_row=5, max_row=4 + NT), Reference(ws, min_col=9, min_row=5, max_row=4 + NT), title="Test rows")
s.marker.symbol = "circle"; s.marker.size = 5; s.graphicalProperties.line.noFill = True
ch.series.append(s)
for col, t in [(20, "Perfect"), (21, "+ tol"), (22, "− tol")]:
    s2 = Series(Reference(ws, min_col=col, min_row=5, max_row=6), Reference(ws, min_col=19, min_row=5, max_row=6), title=t)
    s2.marker.symbol = "none"
    if col != 20:
        s2.graphicalProperties.line.dashStyle = "dash"
    ch.series.append(s2)
ch.x_axis.scaling.min = 700; ch.y_axis.scaling.min = 700
ch.x_axis.delete = False; ch.y_axis.delete = False
ch.height, ch.width = 9, 16
ws.add_chart(ch, "A31")
lc = LineChart(); lc.title = "Actual and predicted power over the test period"; lc.style = 12
lc.add_data(Reference(ws, min_col=9, max_col=10, min_row=4, max_row=4 + NT), titles_from_data=True)
lc.set_categories(Reference(ws, min_col=8, min_row=5, max_row=4 + NT))
lc.y_axis.title = "kW"; lc.y_axis.scaling.min = 700; lc.x_axis.number_format = "dd-mmm hh:mm"
lc.x_axis.delete = False; lc.y_axis.delete = False
lc.height, lc.width = 9, 16
ws.add_chart(lc, "A50")
bc = BarChart(); bc.type = "col"; bc.grouping = "stacked"; bc.overlap = 100; bc.title = "Error per row (Actual − Predicted)"
bc.add_data(Reference(ws, min_col=16, max_col=17, min_row=4, max_row=4 + NT), titles_from_data=True)
bc.set_categories(Reference(ws, min_col=7, min_row=5, max_row=4 + NT))
bc.series[0].graphicalProperties.solidFill = "ED7D31"; bc.series[1].graphicalProperties.solidFill = "5B9BD5"
bc.y_axis.title = "kW"; bc.x_axis.delete = False; bc.y_axis.delete = False
bc.height, bc.width = 9, 16
ws.add_chart(bc, "A69")

# =====================================================================
# Live_Data
# =====================================================================
ws = sheet("Live_Data", {"A": 17, "B": 11, "C": 11, "D": 10, "E": 10, "F": 10, "G": 7, "H": 10, "I": 11, "J": 11, "K": 10,
                         "L": 11, "M": 11, "N": 10, "O": 10, "P": 70})
put(ws, "A1", "Live data – apply the trained constants to the latest readings", F_TITLE)
put(ws, "A2", "Yellow columns = live PV. Stand-in: real 1-minute readings from 2 Feb 2026 (no live feed yet). On site these come from the Historian 'Live' table or OPC-UA every 15–30 min.", F_I)
header(ws, 4, ["Timestamp", "CL1.PV (kW)", "EF.PV2 (kW/RT)", "TIT1 (°C)", "TIT3 (°C)", "TIT4 (°C)", "Valid?", "RT", "Actual kW",
               "Model kW (now)", "Model vs actual", "Best TIT1 (°C)", "Model kW (best)", "Saving (kW)", "kW/RT after", "Advice for operator"])
for i, (_, rec) in enumerate(live.iterrows()):
    r = 5 + i
    for col, v, fm in [("A", rec.DateTime.to_pydatetime(), "yyyy-mm-dd hh:mm"), ("B", round(float(rec.CL1), 1), "#,##0.0"),
                       ("C", round(float(rec.EF2), 4), "0.0000"), ("D", round(float(rec.TIT1), 2), "0.00"),
                       ("E", round(float(rec.TIT3), 2), "0.00"), ("F", round(float(rec.TIT4), 2), "0.00")]:
        put(ws, f"{col}{r}", v, F_IN, FILL_IN, fm)
    f = {"G": f'=IF(COUNT(B{r}:F{r})<5,"No",IF(AND(B{r}>MinCL,F{r}-E{r}>MinDT,E{r}>=CHWS_Lo,E{r}<=CHWS_Hi,D{r}>=CWS_Lo,D{r}<=CWS_Hi,C{r}>=KWRT_Lo,C{r}<=KWRT_Hi),"Yes","No"))',
         "H": f'=IF(G{r}<>"Yes","",B{r}/RT_kW)', "I": f'=IF(G{r}<>"Yes","",C{r}*H{r})',
         "J": f'=IF(G{r}<>"Yes","",Coef_Int+Coef_RT*H{r}+Coef_TIT1*D{r})', "K": f'=IF(G{r}<>"Yes","",(I{r}-J{r})/J{r})',
         "L": f'=IF(G{r}<>"Yes","",IF(Coef_TIT1>0,MIN(D{r},SP_Min),SP_Max))',
         "M": f'=IF(G{r}<>"Yes","",Coef_Int+Coef_RT*H{r}+Coef_TIT1*L{r})', "N": f'=IF(G{r}<>"Yes","",MAX(0,J{r}-M{r}))',
         "O": f'=IF(G{r}<>"Yes","",M{r}/H{r})',
         "P": f'=IF(G{r}<>"Yes","Bad or missing reading – no advice",IF(N{r}<1,"Keep current condenser water temperature","Lower condenser water supply (cooling towers) "&TEXT(D{r},"0.0")&" → "&TEXT(L{r},"0.0")&" °C: save ~"&TEXT(N{r},"0")&" kW"))'}
    for col, v in f.items():
        put(ws, f"{col}{r}", v, F_B if col in "NP" else F_N, FILL_REC if col in "LNP" else None)
    for col, fm in zip("HIJKLMNO", ["#,##0.0", "#,##0.0", "#,##0.0", "+0.0%;-0.0%", "0.0", "#,##0.0", "0.0", "0.0000"]):
        ws[f"{col}{r}"].number_format = fm
LL = 4 + len(live)
put(ws, f"A{LL+2}", "Averages", F_B); put(ws, f"N{LL+2}", f"=AVERAGE(N5:N{LL})", F_B, FILL_KEY, "0.0")
put(ws, f"K{LL+2}", f"=AVERAGE(K5:K{LL})", F_B, FILL_KEY, "+0.0%;-0.0%")
put(ws, f"A{LL+3}", "Model vs actual = how far the real kW is from the model. Drift beyond ±10% for several days → retrain.", F_I)
put(ws, f"A{LL+4}", "Saving ignores the extra cooling tower fan power needed to reach a lower TIT1 – the net saving is smaller (Assumption G9).", F_I)
ws.freeze_panes = "B5"

# =====================================================================
# Optimisation
# =====================================================================
ws = sheet("Optimisation", {"A": 34, "B": 20, "C": 14, "D": 15, "E": 17, "F": 16, "G": 18, "H": 10})
put(ws, "A1", "Optimisation – which condenser water temperature gives the lowest power right now?", F_TITLE)
put(ws, "A2", "Pick a live snapshot (yellow). The model tries every allowed TIT1 (condenser water supply) and picks the lowest kW.", F_I)
put(ws, "A4", "1. Current state", F_H2)
put(ws, "A5", f"Live row to analyse (1–{len(live)})"); put(ws, "B5", 3, F_IN, FILL_IN)
put(ws, "C5", f'=IF(INDEX(Live_Data!$G$5:$G${LL},$B$5)="Yes","Reading valid","INVALID reading – pick another row")', F_I)
cur = [("Timestamp", f"=INDEX(Live_Data!$A$5:$A${LL},$B$5)", "yyyy-mm-dd hh:mm"), ("Cooling load RT", f"=INDEX(Live_Data!$H$5:$H${LL},$B$5)", "#,##0.0"),
       ("TIT1 now (°C)", f"=INDEX(Live_Data!$D$5:$D${LL},$B$5)", "0.00"), ("Actual power (kW)", f"=INDEX(Live_Data!$I$5:$I${LL},$B$5)", "#,##0.0"),
       ("Model kW at current TIT1", "=Coef_Int+Coef_RT*B7+Coef_TIT1*B8", "#,##0.0"), ("Model kW/RT now", "=B10/B7", "0.0000")]
for i, (t, fx, fm) in enumerate(cur):
    put(ws, f"A{6+i}", t); put(ws, f"B{6+i}", fx, fmt=fm, align=Alignment(horizontal="left"))
put(ws, "A14", "2. Try every allowed condenser water supply temperature", F_H2)
header(ws, 15, ["Option", "TIT1 target (°C)", "Predicted kW", "Predicted kW/RT", "Saving vs now (kW)", "Saving ± (95%)", "Prediction ± (95%)", "Best?"])
for k in range(10):
    r = 16 + k
    f = {"A": f'=IF(B{r}="","",{k+1})', "B": f'=IF(SP_Min+{k}*SP_Step<=SP_Max+0.0001,SP_Min+{k}*SP_Step,"")',
         "C": f'=IF(B{r}="","",Coef_Int+Coef_RT*$B$7+Coef_TIT1*B{r})', "D": f'=IF(B{r}="","",C{r}/$B$7)',
         "E": f'=IF(B{r}="","",$B$10-C{r})', "F": f'=IF(B{r}="","",T_Crit*SE_TIT1*ABS(B{r}-$B$8))',
         "G": f'=IF(B{r}="","",T_Crit*SEy)', "H": f'=IF(B{r}="","",IF(C{r}=MIN($C$16:$C$25),"◄ BEST",""))'}
    for col, v in f.items():
        put(ws, f"{col}{r}", v, border=True)
    for col, fm in zip("BCDEFG", ["0.0", "#,##0.0", "0.0000", "+#,##0.0;-#,##0.0", "± 0.0", "± 0.0"]):
        ws[f"{col}{r}"].number_format = fm
ws.conditional_formatting.add("A16:H25", FormulaRule(formula=['$H16="◄ BEST"'], fill=FILL_REC, font=Font(bold=True)))
put(ws, "A26", "Negative saving = that option uses MORE power than now. Saving ± uses the standard error of b2 (only TIT1 changes). Prediction ± is the approximate range of a single prediction.", F_I)
put(ws, "A28", "3. Recommendation", F_H2)
rec = [("Best TIT1 target (°C)", "=INDEX(B16:B25,MATCH(MIN(C16:C25),C16:C25,0))", "0.0"), ("Predicted kW at best", "=MIN(C16:C25)", "#,##0.0"),
       ("Saving (kW)", "=B10-B30", "#,##0.0"), ("Saving uncertainty ± (kW, 95%)", "=T_Crit*SE_TIT1*ABS(B29-B8)", "0.0"),
       ("Saving (%)", "=B31/B10", "0.0%"), ("kW/RT now → after", '=TEXT(B11,"0.0000")&" → "&TEXT(B30/B7,"0.0000")', "@"),
       ("If held all year (MWh/yr)", "=B31*Hours_Yr/1000", "#,##0"), ("If held all year (SGD/yr)", "=B31*Hours_Yr*Tariff", "#,##0")]
for i, (t, fx, fm) in enumerate(rec):
    put(ws, f"A{29+i}", t); put(ws, f"B{29+i}", fx, F_B, FILL_REC, fm, Alignment(horizontal="left"))
put(ws, "A38", "Advice to operator", F_B)
put(ws, "B38", '=IF(B31<1,"Keep current condenser water temperature – already near the lower limit","Lower the condenser water supply (more cooling tower capacity) from "&TEXT(B8,"0.0")&" to "&TEXT(B29,"0.0")&" °C → about "&TEXT(B31,"0")&" kW less chiller power ("&TEXT(B33,"0.0%")&")")', F_B, FILL_REC)
ws.merge_cells("B38:H38")
put(ws, "A40", "Annual figures assume the same load and saving all year – an upper-bound illustration, not a forecast. Tariff is a placeholder (Settings).", F_I)
put(ws, "A41", "Why the best TIT1 is always the lowest limit: the model is linear in TIT1 and b2 > 0, so kW keeps falling as TIT1 falls. The real limit is the outdoor wet-bulb temperature plus the tower approach (SP_Min).", F_I)
put(ws, "A42", "Not included yet: extra cooling tower fan power to reach a lower TIT1. The net plant saving is smaller – needs fan power data (Assumption G9).", F_I)

# =====================================================================
# Real_World_Guide
# =====================================================================
ws = sheet("Real_World_Guide", {"A": 5, "B": 30, "C": 120})
put(ws, "A1", "From this real-data prototype to the delivered solution", F_TITLE)
put(ws, "A3", "A. What changes and what stays", F_H2)
header(ws, 4, ["", "Sheet", "Now (10 days of client data) → Full project"])
guide = [("Raw_Data", "5 × 2 days of 1-minute readings (14,400 rows) → full 3 years at 1 minute ≈ 1.6 million rows, more than Excel's 1,048,576-row limit: use Python for the full data (or 30-minute averages ≈ 52,600 rows in Excel)."),
         ("Settings", "Placeholder limits → real lowest condenser water temperature (plant engineer), real tariff, datasheet rated power."),
         ("Cleaning", "Same 5 checks. Expect more removals in 3-year data (shutdowns, sensor faults) – review the summary."),
         ("Clean_Data", f"Same formulas – extend them down to the number of raw rows (now {N})."),
         ("MLR_Model", "Same LINEST – coefficients re-estimated automatically. Check signs, p-values and VIF again."),
         ("Validation", "Same metrics. If FAIL: check cleaning, add RT², RT × TIT1, or a wet-bulb input."),
         ("Live_Data", "Stand-in 1-minute readings → live query (Historian 'Live' table) or OPC-UA, every 15–30 min."),
         ("Optimisation", "Same logic. Add cooling tower fan power to get the NET saving.")]
for i, (s_, t) in enumerate(guide):
    put(ws, f"B{5+i}", s_, F_B, border=True); put(ws, f"C{5+i}", t, border=True, align=WRAP)
put(ws, "A14", "B. Step by step", F_H2)
steps = [("Get the full data", "Request 3 years of MT Chiller 1 history (CL1, EF.PV, EF.PV2, FIT1, FIT2, HR1, TIT1–TIT4, RunS) at 1-minute resolution (Python) or 30-minute averages (Excel)."),
         ("Confirm the open points", "Assumption sheet: items marked 'Confirm with client' – especially the lowest achievable TIT1 and the power meter (IMSValue)."),
         ("Load into Excel / Python", "Paste into Raw_Data in the same column order, or read with pandas (read_hist.py)."),
         ("Size the formulas", f"Copy the formula rows in Cleaning and Clean_Data down to the new row count, and change ${LAST} ranges to the new last row (Find & Replace '{LAST}')."),
         ("Check the model", "Cleaning summary → MLR_Model (b1, b2 positive? p < 0.05? VIF < 5?) → Validation PASS?"),
         ("Add cooling tower fan power", "Ask for tower fan kW tags; subtract the extra fan power from the chiller saving."),
         ("Repeat per chiller", "One model per chiller (MT CH01–CH05, LT, ML). Same structure, own coefficients."),
         ("Move to Python", "Each sheet = one function: load_history(), clean(), fit(), validate(), recommend(). Client requirement: Python/PowerShell.")]
for i, (t, d_) in enumerate(steps):
    r = 15 + i
    put(ws, f"A{r}", i + 1, F_B); put(ws, f"B{r}", t, F_B); put(ws, f"C{r}", d_, align=WRAP)
put(ws, "A24", "C. Cautions", F_H2)
for i, t in enumerate([f"The model is only valid inside the range it was trained on ({RANGE_TXT}). Do not advise far outside it.",
                       "10 days of data covers only February and August. Results may change with the full 3 years.",
                       "Retrain every 3–6 months, or when Live_Data 'Model vs actual' drifts more than ±10% for several days.",
                       "Chiller-only kW/RT. Whole-plant kW/RT also needs pump and cooling tower fan kW.",
                       "Advice only – the operator always decides. Never advise on rows that fail the validity check."]):
    put(ws, f"A{25+i}", "•", F_B); put(ws, f"B{25+i}", t)

wb.save(OUT)
print("saved", OUT, "rows", N, "live", len(live))

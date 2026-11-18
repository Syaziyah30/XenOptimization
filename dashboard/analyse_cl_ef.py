"""Correlation and day/night checks for MT Chiller 1: CL1.PV and EF.PV2 against the temperature tags."""
import sys
import pandas as pd

EXCEL = r"D:\XenOptimization\20261016_benny - Copy.xlsx"
SHEET = sys.argv[1] if len(sys.argv) > 1 else "2.Hist Value1"   # e.g. "2.Hist Value3"
P = "C10A1_1_C101C_MTCHW_CH01_"
CL, EF = P + "CL1.PV", P + "EF.PV2"
TIT = {"TIT1 (CW supply)": P + "TIT1.PV", "TIT2 (CW return)": P + "TIT2.PV",
       "TIT3 (CHW supply)": P + "TIT3.PV", "TIT4 (CHW return)": P + "TIT4.PV"}
CL_REF, EF_REF = 7741.627441, 0.410726458          # from data/tag_reference.csv

d = pd.read_excel(EXCEL, SHEET).set_index("DateTime")
print(f"{SHEET}: {d.index.min()} to {d.index.max()}, {len(d)} rows\n")

# 1. Correlation of EF.PV2 with each variable (1-minute data)
print("Correlation with EF.PV2 (r):")
print(f"  CL1.PV (1-min)        {d[CL].corr(d[EF]):.2f}")
h = d[[CL, EF]].resample("1h").mean()
print(f"  CL1.PV (hourly avg)   {h[CL].corr(h[EF]):.2f}")
for name, col in TIT.items():
    print(f"  {name:<21} {d[col].corr(d[EF]):.2f}")

# 2. Night (22:00-06:00) vs day (09:00-17:00) averages
night, day = d.between_time("22:00", "06:00"), d.between_time("09:00", "17:00")
print("\nNight vs day average:")
for name, col in [("CL1.PV (kW)", CL), ("EF.PV2 (kW/RT)", EF)] + list(TIT.items()):
    print(f"  {name:<21} night {night[col].mean():9.3f}   day {day[col].mean():9.3f}")

# 3. Share of time above the reference value
print("\nTime above reference:")
print(f"  CL1.PV > {CL_REF:,.1f} kW     {(d[CL] > CL_REF).mean():.0%}")
print(f"  EF.PV2 > {EF_REF:.4f} kW/RT  {(d[EF] > EF_REF).mean():.0%}")

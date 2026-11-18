"""Build CL_Dashboard_Presentation.pdf: slides that explain the CL_Dashboard findings,
each with a 'What to say' script box. Numbers come from CL_Correlation_Hist1-5.xlsx.
Run:  python build_cl_presentation.py
"""
import os
import tempfile

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, Table, TableStyle

HERE = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(HERE, "CL_Correlation_Hist1-5.xlsx")
OUT = os.path.join(HERE, "CL_Dashboard_Presentation.pdf")
IMG = tempfile.mkdtemp(prefix="clpres_")

W, H = 960, 540
NAVY, BLUE, ORANGE, GREY = "#1F3864", "#2E75B6", "#ED7D31", "#7F7F7F"
LIGHT = "#F2F5FA"

pdfmetrics.registerFont(TTFont("Arial", r"C:\Windows\Fonts\arial.ttf"))
pdfmetrics.registerFont(TTFont("Arial-Bold", r"C:\Windows\Fonts\arialbd.ttf"))
pdfmetrics.registerFont(TTFont("Arial-Italic", r"C:\Windows\Fonts\ariali.ttf"))
from reportlab.pdfbase.pdfmetrics import registerFontFamily
registerFontFamily("Arial", normal="Arial", bold="Arial-Bold", italic="Arial-Italic", boldItalic="Arial-Bold")

BODY = ParagraphStyle("body", fontName="Arial", fontSize=13, leading=18, textColor=colors.HexColor("#262626"),
                      alignment=TA_LEFT, spaceAfter=6)
SMALL = ParagraphStyle("small", parent=BODY, fontSize=10.5, leading=14)
SCRIPT = ParagraphStyle("script", parent=BODY, fontSize=10.5, leading=14, textColor=colors.HexColor("#1F1F1F"))
CELL = ParagraphStyle("cell", parent=BODY, fontSize=10.5, leading=13, spaceAfter=0)
CELLB = ParagraphStyle("cellb", parent=CELL, fontName="Arial-Bold", textColor=colors.white)

PERIODS = ["Hist_Value1", "Hist_Value2", "Hist_Value3", "Hist_Value4", "Hist_Value5"]
PLABEL = ["Hist1\nFeb 2024", "Hist2\nAug 2024", "Hist3\nFeb 2025", "Hist4\nAug 2025", "Hist5\nFeb 2026"]

# ------------------------------------------------------------------ data
d1 = pd.read_excel(XLSX, sheet_name="Data 1min")
d30 = pd.read_excel(XLSX, sheet_name="Data 30min")
col = {c.split(" ")[0]: c for c in d1.columns if c not in ("Period", "DateTime")}
CL = col["CL1.PV"]
avg = d1.groupby("Period")[CL].mean().reindex(PERIODS)
overall = d1[CL].mean()
readings = int(d1[CL].notna().sum())

VARS = ["EF.PV", "TIT4.PV", "TIT2.PV", "TIT1.PV", "EF.PV2", "FIT1.PV", "FIT2.PV", "TIT3.PV"]
r_per = {v: [d30[d30.Period == p][CL].corr(d30[d30.Period == p][col[v]]) for p in PERIODS] for v in VARS}
r_pool = {v: d30[CL].corr(d30[col[v]]) for v in VARS}

ok = d1[[CL, col["FIT2.PV"], col["TIT3.PV"], col["TIT4.PV"]]].dropna()
calc = ok[col["FIT2.PV"]] * (ok[col["TIT4.PV"]] - ok[col["TIT3.PV"]]) * 1.163
err = ((calc - ok[CL]).abs() / ok[CL] * 100)

plt.rcParams.update({"font.family": "Arial", "font.size": 11, "axes.spines.top": False, "axes.spines.right": False})


def save(fig, name):
    p = os.path.join(IMG, name)
    fig.savefig(p, dpi=200, bbox_inches="tight")
    plt.close(fig)
    return p


# chart: trend per period
fig, axes = plt.subplots(1, 5, figsize=(11, 3.2), sharey=True)
for ax, p, lab in zip(axes, PERIODS, PLABEL):
    s = d30[d30.Period == p]
    hrs = np.arange(len(s)) * 0.5
    ax.plot(hrs, s[CL], color=BLUE, lw=1.6)
    ax.axhline(overall, color=ORANGE, ls="--", lw=1.2)
    ax.set_title(lab.replace("\n", "  "), fontsize=10.5)
    ax.set_xticks([0, 24, 48])
    ax.set_xlabel("hours", fontsize=9)
axes[0].set_ylabel("CL1.PV (kW)")
axes[0].plot([], [], color=BLUE, label="CL1.PV (30-min)")
axes[0].plot([], [], color=ORANGE, ls="--", label=f"overall average {overall:,.0f} kW")
fig.legend(loc="upper center", ncol=2, frameon=False, bbox_to_anchor=(0.5, 1.08))
fig.tight_layout()
IMG_TREND = save(fig, "trend.png")

# chart: average by period
fig, ax = plt.subplots(figsize=(6.2, 3.4))
x = np.arange(5)
ax.plot(x, avg.values, marker="o", color=BLUE, lw=2)
ax.axhline(overall, color=ORANGE, ls="--", lw=1.4)
for i, v in enumerate(avg.values):
    ax.text(i, v + 25, f"{v:,.0f}", ha="center", fontsize=10)
ax.text(-0.2, overall - 50, f"overall {overall:,.0f}", color=ORANGE, ha="left", fontsize=10)
ax.set_xlim(-0.3, 4.3)
ax.set_xticks(x, PLABEL, fontsize=9.5)
ax.set_ylabel("Average CL1.PV (kW)")
ax.set_ylim(7500, 8450)
fig.tight_layout()
IMG_AVG = save(fig, "avg.png")

# chart: correlation dot plot
ORDER = ["EF.PV", "EF.PV2", "FIT1.PV", "FIT2.PV", "TIT1.PV", "TIT2.PV", "TIT3.PV", "TIT4.PV"]  # top to bottom
order = ORDER[::-1]  # matplotlib draws index 0 at the bottom
fig, ax = plt.subplots(figsize=(6.4, 3.9))
for band, c in [(0.8, "#DCE9F7"), (0.6, "#E9F1FA"), (0.4, "#F4F8FC")]:
    ax.axvspan(band, 1.0 if band == 0.8 else band + 0.2, color=c, zorder=0)
ax.axvline(0, color="#BFBFBF", lw=1)
for i, v in enumerate(order):
    rs = r_per[v]
    ax.plot([min(rs), max(rs)], [i, i], color="#BFBFBF", lw=5, solid_capstyle="round", zorder=1)
    ax.scatter(rs, [i] * 5, color=GREY, s=18, zorder=2)
    ax.scatter([r_pool[v]], [i], color=ORANGE, s=70, marker="D", zorder=3)
ax.set_yticks(range(len(order)), order)
ax.set_xlim(-1, 1)
ax.set_xlabel("Pearson r with CL1.PV (30-min data)")
ax.scatter([], [], color=GREY, s=18, label="each period (Hist1–5)")
ax.scatter([], [], color=ORANGE, s=60, marker="D", label="all periods pooled")
ax.legend(loc="upper left", fontsize=9, frameon=False)
fig.tight_layout()
IMG_CORR = save(fig, "corr.png")

# chart: formula check
fig, ax = plt.subplots(figsize=(4.6, 3.6))
smp = ok.sample(3000, random_state=1)
cs = calc.loc[smp.index]
ax.scatter(smp[CL], cs, s=4, alpha=0.35, color=BLUE)
lo, hi = 7300, 8800
ax.plot([lo, hi], [lo, hi], color=ORANGE, lw=1.4, ls="--")
ax.set_xlim(lo, hi)
ax.set_ylim(lo, hi)
ax.set_xlabel("CL1.PV recorded by SCADA (kW)")
ax.set_ylabel("Calculated by formula (kW)")
fig.tight_layout()
IMG_FORM = save(fig, "formula.png")


# ------------------------------------------------------------------ slide helpers
def para(c, text, x, y_top, w, style=BODY):
    p = Paragraph(text, style)
    _, h = p.wrap(w, 1000)
    p.drawOn(c, x, y_top - h)
    return y_top - h - style.spaceAfter


def bullets(c, items, x, y_top, w, style=BODY):
    for it in items:
        y_top = para(c, f"•&nbsp;&nbsp;{it}", x, y_top, w, style)
    return y_top


def image(c, path, x, y, w=None, h=None):
    from reportlab.lib.utils import ImageReader
    ir = ImageReader(path)
    iw, ih = ir.getSize()
    if w and not h:
        h = w * ih / iw
    elif h and not w:
        w = h * iw / ih
    c.drawImage(ir, x, y, w, h)
    return w, h


def table(c, data, x, y_top, widths, header=True, font=10.5):
    st = ParagraphStyle("t", parent=CELL, fontSize=font, leading=font + 2.5)
    sth = ParagraphStyle("th", parent=CELLB, fontSize=font, leading=font + 2.5)
    rows = [[Paragraph(str(v), sth if (header and i == 0) else st) for v in r] for i, r in enumerate(data)]
    t = Table(rows, colWidths=widths)
    style = [("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D9D9D9")),
             ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
             ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]
    if header:
        style.append(("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(NAVY)))
    t.setStyle(TableStyle(style))
    _, h = t.wrap(sum(widths), 1000)
    t.drawOn(c, x, y_top - h)
    return y_top - h


SCRIPT_TOP = 128


class Deck:
    def __init__(self):
        self.c = canvas.Canvas(OUT, pagesize=(W, H))
        self.c.setTitle("MT Chiller 1 – Cooling Load (CL) Dashboard findings")
        self.n = 0

    def slide(self, title, page_ref, script):
        c = self.c
        if self.n:
            c.showPage()
        self.n += 1
        c.setFillColor(colors.white)
        c.rect(0, 0, W, H, stroke=0, fill=1)
        c.setFillColor(colors.HexColor(NAVY))
        c.rect(0, H - 62, W, 62, stroke=0, fill=1)
        c.setFillColor(colors.white)
        c.setFont("Arial-Bold", 22)
        c.drawString(30, H - 40, title)
        if page_ref:
            c.setFont("Arial", 10.5)
            c.drawRightString(W - 30, H - 38, page_ref)
        # script box
        c.setFillColor(colors.HexColor(LIGHT))
        c.setStrokeColor(colors.HexColor("#C9D6EA"))
        c.roundRect(20, 22, W - 40, SCRIPT_TOP - 30, 6, stroke=1, fill=1)
        c.setFillColor(colors.HexColor(BLUE))
        c.setFont("Arial-Bold", 10)
        c.drawString(34, SCRIPT_TOP - 24, "WHAT TO SAY")
        para(c, script, 34, SCRIPT_TOP - 30, W - 68, SCRIPT)
        c.setFillColor(colors.HexColor(GREY))
        c.setFont("Arial", 8.5)
        c.drawString(20, 8, "MT Chiller 1 – CL Dashboard findings")
        c.drawRightString(W - 20, 8, str(self.n))
        return c

    def save(self):
        self.c.save()


TOP = H - 82  # first content line
dk = Deck()

# 1 ---------------------------------------------------------------- title
c = dk.slide("MT Chiller 1 – What the Cooling Load (CL) data tells us", "",
             "Today I'll walk you through what we learned from the CL Dashboard in Power BI. It covers five "
             "two-day periods of 1-minute data from MT Chiller 1. By the end, you'll see why our model will "
             "predict <b>chiller power</b>, not cooling load, and which setting we will recommend the operator "
             "to adjust.")
c.setFillColor(colors.HexColor(NAVY))
c.setFont("Arial-Bold", 30)
c.drawString(60, 380, "Cooling Load analysis – CL Dashboard")
c.setFont("Arial", 16)
c.setFillColor(colors.HexColor("#404040"))
c.drawString(60, 345, "Chiller energy optimisation, Phase 1 (chillers only, setpoint recommendation)")
c.drawString(60, 318, "Data: MT Chiller 1, Hist_Value1–5 (Feb/Aug 2024–2026), 1-minute readings")
c.setFillColor(colors.HexColor(ORANGE))
c.rect(60, 300, 120, 4, stroke=0, fill=1)
c.setFillColor(colors.HexColor(GREY))
c.setFont("Arial", 12)
c.drawString(60, 270, "Source: CL_Dashboard.pbip  |  CL_Correlation_Hist1-5.xlsx")

# 2 ---------------------------------------------------------------- purpose
c = dk.slide("Why we did this analysis", "",
             "The client wants to cut chiller energy. They want us to <b>recommend setpoints</b>, and their "
             "operator decides. Before building any model, we need to understand the data: what moves, what "
             "moves together, and what we can actually change. That's what the CL Dashboard is for. It's step 1 "
             "and step 2; the model is step 3.")
y = para(c, "<b>Client need:</b> lower chiller energy by recommending better setpoints. The operator stays in control.",
         40, TOP, 880)
table(c, [["Step", "Question", "Tool (dashboard page)", "Output"],
          ["1. Trend", "How does the load behave? Is the data good?", "CL Combined Dashboard, All Variables, Average Variable",
           "Patterns, data quality, normal values"],
          ["2. Correlation", "Which variables move with the load, and how strongly?", "CL Correlation Scatter",
           "Shortlist of variables for the model"],
          ["3. MLR model", "How much power can we save by changing a setting?", "ExampleProject.xlsx (next phase)",
           "Recommended setpoint + kW saving"]],
      40, y - 12, [110, 290, 270, 210], font=12)

# 3 ---------------------------------------------------------------- data
c = dk.slide("The data we used", "Dashboard page: All Variables",
             f"We used the client's 1-minute data for MT Chiller 1: five periods of two days each, February and August "
             f"of 2024 to 2026. That's {readings:,} readings per tag out of 14,400; only one reading is blank, so the data is "
             "complete and we can trust it. One limitation: each period is only two days, so I'll ask the client "
             "for a longer history.")
table(c, [["Period", "Dates", "Readings"]] +
      [[p, f"{d1[d1.Period == p].DateTime.min():%d %b %Y} – {d1[d1.Period == p].DateTime.max():%d %b %Y}", "2,880 (1-min)"]
       for p in PERIODS], 40, TOP, [110, 200, 110], font=11)
table(c, [["Tag", "Meaning", "Unit"],
          ["CL1.PV", "Cooling load (heat removed)", "kW"], ["EF.PV", "Chiller loading", "%"],
          ["EF.PV2", "Efficiency (lower = better)", "kW/RT"], ["TIT1 / TIT2", "Condenser water supply / return temp", "°C"],
          ["TIT3 / TIT4", "Chilled water supply / return temp", "°C"], ["FIT1 / FIT2", "Condenser / chilled water flow", "m³/h"]],
      490, TOP, [90, 270, 70], font=11)

# 4 ---------------------------------------------------------------- formula
c = dk.slide("What is Cooling Load (CL)?", "",
             "CL is the heat the factory sends to the chiller. It's calculated from the chilled water: how much "
             "water flows, times how much it warmed up in the factory, times 1.163, a fixed number for water. "
             f"I checked this against every reading: the formula matches the SCADA value within {err.mean():.2f}% "
             "on average. This matters later: TIT4 and FIT2 are <b>part of</b> CL, so they'll naturally look linked to it.")
y = para(c, "<font size=17><b>CL1.PV (kW) = FIT2 × (TIT4 − TIT3) × 1.163</b></font>", 40, TOP - 10, 520)
bullets(c, ["<b>FIT2</b>: chilled water flow (m³/h)",
            "<b>TIT4 − TIT3</b>: how much the water warmed up in the factory (°C)",
            "<b>1.163</b>: heat carried by 1 m³/h of water per 1 °C (4.186 × 1000 ÷ 3600)",
            f"<b>Check:</b> {len(err):,} readings, average difference <b>{err.mean():.2f}%</b>, largest {err.max():.1f}%"],
        40, y - 14, 500)
image(c, IMG_FORM, 590, SCRIPT_TOP + 14, w=330)

# 5 ---------------------------------------------------------------- trend
hi_p, lo_p = avg.idxmax(), avg.idxmin()
c = dk.slide("Step 1 – How the cooling load behaves", "Dashboard page: CL Combined Dashboard",
             f"This is the cooling load over each two-day period. The orange line is the overall average, "
             f"{overall:,.0f} kW. The load moves between about {d30[CL].min():,.0f} and {d30[CL].max():,.0f} kW. When I tick "
             "other variables on the dashboard, <b>TIT4 follows CL almost line for line</b>; chiller loading (EF.PV) and the "
             "condenser water temperatures (TIT1, TIT2) rise and fall with it; <b>TIT3 stays flat</b> at about 12.8 °C "
             "because it's a setpoint; the flows (FIT1, FIT2) show no consistent pattern.")
image(c, IMG_TREND, 30, SCRIPT_TOP + 10, w=W - 60)

# 6 ---------------------------------------------------------------- averages
c = dk.slide("Step 1 – Average load per period", "Dashboard page: Average Variable",
             f"Here each dot is the average load of one period against the overall average of {overall:,.0f} kW. "
             f"The highest is {hi_p.replace('_Value', '')}, August 2025, at {avg.max():,.0f} kW; the lowest is "
             f"{lo_p.replace('_Value', '')}, February 2024, at {avg.min():,.0f} kW. August periods are above average, and "
             "the February load has also been rising year by year. These averages give us the <b>normal operating values</b> "
             "to compare our recommendations against.")
image(c, IMG_AVG, 30, SCRIPT_TOP + 10, w=505)
sm = d1.groupby("Period")[[col["TIT1.PV"], col["EF.PV2"], col["EF.PV"]]].mean().reindex(PERIODS)
rows = [["Period", "CL (kW)", "TIT1 (°C)", "EF.PV2 (kW/RT)", "EF.PV (%)"]]
for p in PERIODS:
    rows.append([p.replace("_Value", ""), f"{avg[p]:,.0f}", f"{sm.loc[p].iloc[0]:.2f}", f"{sm.loc[p].iloc[1]:.3f}",
                 f"{sm.loc[p].iloc[2]:.1f}"])
rows.append(["<b>Overall</b>", f"<b>{overall:,.0f}</b>", f"<b>{d1[col['TIT1.PV']].mean():.2f}</b>",
             f"<b>{d1[col['EF.PV2']].mean():.3f}</b>", f"<b>{d1[col['EF.PV']].mean():.1f}</b>"])
table(c, rows, 555, TOP - 20, [70, 70, 70, 90, 70], font=10.5)
para(c, "Best efficiency was Hist3 (Feb 2025): lowest TIT1 and lowest kW/RT.", 555, TOP - 205, 370, SMALL)

# 7 ---------------------------------------------------------------- correlation
c = dk.slide("Step 2 – Which variables move with the load?", "Dashboard page: CL Correlation Scatter",
             "Pearson r measures how closely two things move together: 1 is perfect, 0 is no link. Grey dots are the five "
             "periods, the orange diamond is all periods together. EF.PV and TIT4 are the strongest and consistent in every "
             "period. TIT2, TIT1 and EF.PV2 are moderate to strong. FIT1 and FIT2 change direction between periods, so we "
             "can't rely on them. TIT3 has no link. We judge a variable by both its <b>range across periods</b> and its "
             "<b>pooled r</b>, not a simple average.")
image(c, IMG_CORR, 30, SCRIPT_TOP + 6, h=H - 62 - SCRIPT_TOP - 14)
rows = [["Variable", "Range (5 periods)", "Pooled r"]]
for v in ORDER:
    rows.append([v, f"{min(r_per[v]):+.2f} to {max(r_per[v]):+.2f}", f"{r_pool[v]:+.2f}"])
table(c, rows, 590, TOP - 10, [90, 150, 80], font=10.5)
para(c, "Strength: |r| ≥ 0.8 very strong, ≥ 0.6 strong, ≥ 0.4 moderate, ≥ 0.2 weak.", 590, TOP - 245, 330, SMALL)

# 8 ---------------------------------------------------------------- meaning
c = dk.slide("Step 2 – What the correlation really means", "",
             "The key point: none of these variables <b>drives</b> the load. TIT4 and FIT2 are part of how CL is calculated. "
             "EF.PV, TIT1, TIT2 and EF.PV2 go up <b>because</b> the load goes up. Think of a restaurant: the number of customers "
             "is the load; counting plates is part of counting customers; the gas bill rises because of the customers. "
             "So the cooling load is the factory's demand. We can't optimise it; we can only serve it with less power.")
table(c, [["Group", "Variables", "Why they're linked to CL"],
          ["Part of the CL formula", "TIT4, FIT2", "CL = FIT2 × (TIT4 − TIT3) × 1.163, so they're linked by definition"],
          ["Result of the load", "EF.PV, TIT2, TIT1, EF.PV2", "More heat to remove → chiller works harder, condenser water gets warmer"],
          ["Not reliable", "FIT1, FIT2", "Direction flips between periods (depends on pump operation)"],
          ["No link", "TIT3", "Fixed chilled water setpoint (~12.8 °C)"]],
      40, TOP, [190, 220, 470], font=12)
c.setFillColor(colors.HexColor("#FFF4EC"))
c.setStrokeColor(colors.HexColor(ORANGE))
c.roundRect(40, SCRIPT_TOP + 14, 880, 52, 6, stroke=1, fill=1)
para(c, "<b>Conclusion:</b> CL is the factory's demand. It's an <b>input</b> to our model, not the thing we optimise.",
     56, SCRIPT_TOP + 56, 850, ParagraphStyle("k", parent=BODY, fontSize=15, leading=19))

# 9 ---------------------------------------------------------------- approach
c = dk.slide("Step 3 – So what will the model do?", "Next: MLR model (ExampleProject.xlsx)",
             "Because the load can't be changed, the model predicts <b>chiller power</b>, with the load as an input and the "
             "condenser water temperature, TIT1, as the setting we can change through the cooling towers. Our draft model "
             "fits well: at the same load, each 1 °C lower TIT1 saves about 30 kW. It passes the ASHRAE test on Feb 2026 data, "
             "though only just on R², so we need the client's longer data and their real power meter to confirm it.")
para(c, "<font size=16><b>Power (kW) = −857.6 + 0.3675 × RT + 29.89 × TIT1</b></font>", 40, TOP - 6, 880)
bullets(c, ["<b>Y</b> = chiller power (kW): the energy we want to save",
            "<b>RT</b> = CL1.PV ÷ 3.517: the load, which we take as given",
            "<b>TIT1</b> = condenser water supply temp: the setting the operator can change",
            "+100 RT of load → +36.8 kW;  <b>1 °C lower TIT1 → about 30 kW saved</b> at the same load"],
        40, TOP - 46, 470)
table(c, [["Check (ASHRAE Guideline 14)", "Limit", "Result"],
          ["R² – training (Hist1–4)", "≥ 0.75", "0.994"],
          ["R² – test (Hist5, Feb 2026)", "≥ 0.75", "0.751"],
          ["CV(RMSE) – test", "≤ 30%", "1.5%"],
          ["NMBE – test", "≤ ±10%", "1.5%"]],
      540, TOP - 46, [200, 70, 110], font=11)
para(c, "Limitations: power is calculated (EF.PV2 × RT), not metered; cooling tower fan power not yet included; "
        "each period is only 2 days.", 540, TOP - 185, 380, SMALL)

# 10 --------------------------------------------------------------- next steps
c = dk.slide("Next steps", "",
             "To finish Phase 1 by 31 December, we need the client's data by 20 November. I'd like to confirm today that the "
             "model target is chiller power, and agree on the data request: a direct power meter tag, the lowest safe TIT1, "
             "cooling tower fan power so we know the net saving is real, and a longer history. If the data comes late, the "
             "end date moves by the same number of days.")
table(c, [["When", "What", "Need from client"],
          ["16 – 20 Nov", "Agree approach, send data request", "Power meter tag, lowest safe TIT1, fan power, longer history"],
          ["23 Nov – 4 Dec", "Final MLR model and validation", "Data from step 1"],
          ["7 – 11 Dec", "Recommendation tool (setpoint + kW saving)", "Min/max setpoint limits"],
          ["14 – 18 Dec", "Internal review and client demo", "Meeting slot"],
          ["21 – 31 Dec", "Trial on live data, summary report", "Operator feedback"],
          ["31 Dec 2026", "Phase 1 complete (target)", "Data by 20 Nov; delays shift the date"]],
      40, TOP, [140, 330, 410], font=12)

dk.save()
print("saved", OUT, "| slides:", dk.n)

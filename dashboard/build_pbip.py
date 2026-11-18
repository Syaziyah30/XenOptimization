"""Generate the Power BI Project (PBIP) for the MT Chiller 1 trend dashboard.

Output (next to this script):
    ChillerTrend.pbip                 <- open this in Power BI Desktop
    ChillerTrend.SemanticModel/       <- TMDL data model (Power Query + DAX)
    ChillerTrend.Report/              <- report pages and visuals

The model reads the historian sheets "2.Hist Value1".."2.Hist Value5" straight
from the Excel file, so after the Excel is updated a Refresh in Power BI is
enough. Reference (set) values live in data/tag_reference.csv.

Re-run:  python build_pbip.py
"""

import csv
import json
import shutil
import uuid
from pathlib import Path

HERE = Path(__file__).resolve().parent
EXCEL_FILE = r"D:\XenOptimization\20261016_benny - Copy.xlsx"
REFERENCE_FILE = str(HERE / "data" / "tag_reference.csv")
NAME = "ChillerTrend"
N_PERIODS = 5
BASE_THEME = "CY24SU10"  # copied from the Power BI Desktop install into data/theme

ACTUAL_COLOR = "#1F77B4"     # matplotlib blue, same as the notebook
REFERENCE_COLOR = "#D62728"  # red reference line

with open(HERE / "data" / "tag_reference.csv", newline="", encoding="utf-8") as f:
    TAGS = list(csv.DictReader(f))

PERIODS = [f"Hist_Value{i}" for i in range(1, N_PERIODS + 1)]


def guid():
    return str(uuid.uuid4())


def write(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def write_json(path, obj):
    write(path, json.dumps(obj, indent=2))


def m_indent(code, tabs):
    pad = "\t" * tabs
    return "\n".join(pad + line if line.strip() else "" for line in code.strip("\n").splitlines())


# --------------------------------------------------------------------------
# Semantic model (TMDL)
# --------------------------------------------------------------------------
TAG_LIST_M = ", ".join(f'"{t["Tag"]}"' for t in TAGS)

TREND_M = f"""
let
    Source = Excel.Workbook(File.Contents(ExcelFilePath), null, true),
    SelectedTags = {{{TAG_LIST_M}}},
    HistSheets = Table.SelectRows(Source, each [Kind] = "Sheet" and Text.StartsWith([Name], "2.Hist Value")),
    AddLong = Table.AddColumn(HistSheets, "Long", each
        let
            Promoted = Table.PromoteHeaders([Data], [PromoteAllScalars = true]),
            Kept = Table.SelectColumns(Promoted, {{"DateTime"}} & SelectedTags, MissingField.UseNull),
            Unpivoted = Table.UnpivotOtherColumns(Kept, {{"DateTime"}}, "Tag", "Value")
        in
            Unpivoted),
    AddPeriod = Table.AddColumn(AddLong, "Period", each "Hist_Value" & Text.AfterDelimiter([Name], "Hist Value"), type text),
    Slim = Table.SelectColumns(AddPeriod, {{"Period", "Long"}}),
    Expanded = Table.ExpandTableColumn(Slim, "Long", {{"DateTime", "Tag", "Value"}}),
    Typed = Table.TransformColumnTypes(Expanded, {{{{"DateTime", type datetime}}, {{"Tag", type text}}, {{"Value", type number}}}}),
    NoNulls = Table.SelectRows(Typed, each [Value] <> null and [DateTime] <> null)
in
    NoNulls
"""

PERIOD_M = """
let
    Source = Trend,
    Grouped = Table.Group(Source, {"Period"}, {
        {"StartDateTime", each List.Min([DateTime]), type datetime},
        {"EndDateTime", each List.Max([DateTime]), type datetime}}),
    AddRange = Table.AddColumn(Grouped, "DateRange", each
        DateTime.ToText([StartDateTime], [Format = "dd MMM yyyy", Culture = "en-US"]) & " - " &
        DateTime.ToText([EndDateTime], [Format = "dd MMM yyyy", Culture = "en-US"]), type text),
    AddTitle = Table.AddColumn(AddRange, "PeriodTitle", each [Period] & "  |  " & [DateRange], type text),
    AddOrder = Table.AddColumn(AddTitle, "PeriodOrder", each Number.FromText(Text.AfterDelimiter([Period], "Hist_Value")), Int64.Type)
in
    AddOrder
"""

TAG_M = """
let
    Source = Csv.Document(File.Contents(ReferenceFilePath), [Delimiter = ",", Columns = 6, Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Promoted, {
        {"SortOrder", Int64.Type}, {"Tag", type text}, {"ShortName", type text},
        {"Description", type text}, {"Unit", type text}, {"ReferenceValue", type number}}, "en-US"),
    AddLabel = Table.AddColumn(Typed, "TagLabel", each
        [ShortName] & " - " & Text.Trim(Text.Replace([Description], "(CUP) Level 1 MT Chiller 1", "")) & " (" & [Unit] & ")", type text)
in
    AddLabel
"""

MEASURES = [
    ("Actual Value", "AVERAGE ( Trend[Value] )", "#,0.000"),
    ("Reference Value",
     "IF ( NOT ISEMPTY ( Trend ), SELECTEDVALUE ( DimTag[ReferenceValue] ) )", "#,0.000"),
    ("Average Value", "AVERAGE ( Trend[Value] )", "#,0.000"),
    ("Min Value", "MIN ( Trend[Value] )", "#,0.000"),
    ("Max Value", "MAX ( Trend[Value] )", "#,0.000"),
    ("Deviation from Reference %",
     "DIVIDE ( [Average Value] - [Reference Value], [Reference Value] )", "0.00%"),
    ("% Time Above Reference",
     "VAR ref = SELECTEDVALUE ( DimTag[ReferenceValue] )\n"
     "RETURN\n"
     "    IF (\n"
     "        NOT ISBLANK ( ref ),\n"
     "        DIVIDE ( CALCULATE ( COUNTROWS ( Trend ), Trend[Value] > ref ), COUNTROWS ( Trend ) )\n"
     "    )", "0.0%"),
    # Shared Y-axis across the 5 period charts (like sharey=True in matplotlib)
    ("Y Axis Min",
     "VAR ref = SELECTEDVALUE ( DimTag[ReferenceValue] )\n"
     "VAR lo = CALCULATE ( MIN ( Trend[Value] ), REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( Trend[Period] ), REMOVEFILTERS ( Trend[DateTime] ) )\n"
     "VAR hi = CALCULATE ( MAX ( Trend[Value] ), REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( Trend[Period] ), REMOVEFILTERS ( Trend[DateTime] ) )\n"
     "VAR mn = IF ( ISBLANK ( ref ), lo, MIN ( lo, ref ) )\n"
     "VAR mx = IF ( ISBLANK ( ref ), hi, MAX ( hi, ref ) )\n"
     "RETURN mn - ( mx - mn ) * 0.05", "#,0.000"),
    ("Y Axis Max",
     "VAR ref = SELECTEDVALUE ( DimTag[ReferenceValue] )\n"
     "VAR lo = CALCULATE ( MIN ( Trend[Value] ), REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( Trend[Period] ), REMOVEFILTERS ( Trend[DateTime] ) )\n"
     "VAR hi = CALCULATE ( MAX ( Trend[Value] ), REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( Trend[Period] ), REMOVEFILTERS ( Trend[DateTime] ) )\n"
     "VAR mn = IF ( ISBLANK ( ref ), lo, MIN ( lo, ref ) )\n"
     "VAR mx = IF ( ISBLANK ( ref ), hi, MAX ( hi, ref ) )\n"
     "RETURN mx + ( mx - mn ) * 0.05", "#,0.000"),
    ("Period Title", 'SELECTEDVALUE ( DimPeriod[PeriodTitle], "All periods" )', None),
    ("Unit Axis Title", 'SELECTEDVALUE ( DimTag[ShortName] ) & " (" & SELECTEDVALUE ( DimTag[Unit] ) & ")"', None),
    ("Selected Tag Title",
     "VAR lbl = SELECTEDVALUE ( DimTag[TagLabel], \"Select a variable\" )\n"
     "VAR tag = SELECTEDVALUE ( DimTag[Tag] )\n"
     "VAR ref = SELECTEDVALUE ( DimTag[ReferenceValue] )\n"
     "VAR unit = SELECTEDVALUE ( DimTag[Unit] )\n"
     "RETURN\n"
     "    \"Trend Graph - \" & tag & \"   (\" & lbl & \")   Reference Value = \" & FORMAT ( ref, \"#,0.00##\" ) & \" \" & unit", None),
]


def tmdl_measure(name, expr, fmt):
    # "Trend" is a reserved word in DAX, so the table name must always be quoted
    expr = expr.replace("Trend[", "'Trend'[").replace("( Trend )", "( 'Trend' )")
    lines = expr.splitlines()
    if len(lines) == 1:
        out = f"\tmeasure '{name}' = {lines[0]}\n"
    else:
        out = f"\tmeasure '{name}' =\n" + "\n".join("\t\t\t" + l for l in lines) + "\n"
    if fmt:
        out += f"\t\tformatString: {fmt}\n"
    out += f"\t\tlineageTag: {guid()}\n\n"
    return out


def tmdl_column(name, dtype, fmt=None, sort_by=None, hidden=False, summarize="none"):
    out = f"\tcolumn {name}\n\t\tdataType: {dtype}\n"
    if fmt:
        out += f"\t\tformatString: {fmt}\n"
    if hidden:
        out += "\t\tisHidden\n"
    out += f"\t\tlineageTag: {guid()}\n\t\tsummarizeBy: {summarize}\n\t\tsourceColumn: {name}\n"
    if sort_by:
        out += f"\t\tsortByColumn: {sort_by}\n"
    out += "\n"
    return out


def tmdl_partition(table, m):
    return (f"\tpartition {table} = m\n\t\tmode: import\n\t\tsource =\n"
            f"{m_indent(m, 4)}\n\n")


def build_model(root):
    d = root / "definition"
    write_json(root / "definition.pbism", {"version": "4.2", "settings": {}})
    write_json(root / ".platform", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "SemanticModel", "displayName": NAME},
        "config": {"version": "2.0", "logicalId": guid()}})

    write(d / "database.tmdl", "database\n\tcompatibilityLevel: 1567\n\n")
    write(d / "model.tmdl",
          "model Model\n"
          "\tculture: en-US\n"
          "\tdefaultPowerBIDataSourceVersion: powerBI_V3\n"
          "\tsourceQueryCulture: en-US\n"
          "\tdataAccessOptions\n"
          "\t\tlegacyRedirects\n"
          "\t\treturnErrorValuesAsNull\n\n"
          "annotation __PBI_TimeIntelligenceEnabled = 0\n\n"
          'annotation PBI_QueryOrder = ["ExcelFilePath","ReferenceFilePath","Trend","DimTag","DimPeriod"]\n\n'
          "ref table Trend\nref table DimTag\nref table DimPeriod\n")

    write(d / "expressions.tmdl",
          f'expression ExcelFilePath = "{EXCEL_FILE}" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]\n'
          f"\tlineageTag: {guid()}\n\n"
          "\tannotation PBI_ResultType = Text\n\n"
          f'expression ReferenceFilePath = "{REFERENCE_FILE}" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]\n'
          f"\tlineageTag: {guid()}\n\n"
          "\tannotation PBI_ResultType = Text\n\n")

    trend = f"table Trend\n\tlineageTag: {guid()}\n\n"
    for name, expr, fmt in MEASURES:
        trend += tmdl_measure(name, expr, fmt)
    trend += tmdl_column("Period", "string")
    trend += tmdl_column("DateTime", "dateTime", fmt="General Date")
    trend += tmdl_column("Tag", "string")
    trend += tmdl_column("Value", "double", fmt="#,0.000")
    trend += tmdl_partition("Trend", TREND_M)
    write(d / "tables" / "Trend.tmdl", trend)

    tag = f"table DimTag\n\tlineageTag: {guid()}\n\n"
    tag += tmdl_column("SortOrder", "int64", fmt="0", hidden=True)
    tag += tmdl_column("Tag", "string", sort_by="SortOrder")
    tag += tmdl_column("ShortName", "string", sort_by="SortOrder")
    tag += tmdl_column("Description", "string")
    tag += tmdl_column("Unit", "string")
    tag += tmdl_column("ReferenceValue", "double", fmt="#,0.000")
    tag += tmdl_column("TagLabel", "string", sort_by="SortOrder")
    tag += tmdl_partition("DimTag", TAG_M)
    write(d / "tables" / "DimTag.tmdl", tag)

    per = f"table DimPeriod\n\tlineageTag: {guid()}\n\n"
    per += tmdl_column("Period", "string", sort_by="PeriodOrder")
    per += tmdl_column("StartDateTime", "dateTime", fmt="General Date")
    per += tmdl_column("EndDateTime", "dateTime", fmt="General Date")
    per += tmdl_column("DateRange", "string", sort_by="PeriodOrder")
    per += tmdl_column("PeriodTitle", "string", sort_by="PeriodOrder")
    per += tmdl_column("PeriodOrder", "int64", fmt="0", hidden=True)
    per += tmdl_partition("DimPeriod", PERIOD_M)
    write(d / "tables" / "DimPeriod.tmdl", per)

    write(d / "relationships.tmdl",
          f"relationship {guid()}\n\tfromColumn: Trend.Tag\n\ttoColumn: DimTag.Tag\n\n"
          f"relationship {guid()}\n\tfromColumn: Trend.Period\n\ttoColumn: DimPeriod.Period\n")


# --------------------------------------------------------------------------
# Report (PBIR-legacy report.json)
# --------------------------------------------------------------------------
def lit(v):
    if isinstance(v, bool):
        s = "true" if v else "false"
    elif isinstance(v, int):
        s = f"{v}L"
    elif isinstance(v, float):
        s = f"{v}D"
    else:
        s = "'" + str(v).replace("'", "''") + "'"
    return {"expr": {"Literal": {"Value": s}}}


def color(hex_):
    return {"solid": {"color": lit(hex_)}}


def measure_expr(entity, prop):
    return {"expr": {"Measure": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}}}


def col_ref(src, prop):
    return {"Column": {"Expression": {"SourceRef": {"Source": src}}, "Property": prop}}


def meas_ref(src, prop):
    return {"Measure": {"Expression": {"SourceRef": {"Source": src}}, "Property": prop}}


ALIAS = {"Trend": "t", "DimTag": "g", "DimPeriod": "p"}


def query(fields, order_by=None):
    """fields: list of (kind, entity, prop) with kind in {'col','meas'}; order_by: (entity, column) ascending."""
    used = []
    for _, e, _ in fields:
        if e not in used:
            used.append(e)
    select = []
    for kind, e, p in fields:
        ref = col_ref(ALIAS[e], p) if kind == "col" else meas_ref(ALIAS[e], p)
        ref["Name"] = f"{e}.{p}"
        ref["NativeReferenceName"] = p
        select.append(ref)
    q = {"Version": 2, "From": [{"Name": ALIAS[e], "Entity": e, "Type": 0} for e in used],
         "Select": select}
    if order_by:
        # Continuous X-axis is only allowed when the chart is sorted by the axis field
        q["OrderBy"] = [{"Direction": 1, "Expression": col_ref(ALIAS[order_by[0]], order_by[1])}]
    return q


def in_filter(entity, prop, values, name=None):
    a = ALIAS[entity]
    return {
        "name": name or uuid.uuid4().hex[:20],
        "expression": {"Column": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}},
        "filter": {"Version": 2, "From": [{"Name": a, "Entity": entity, "Type": 0}],
                   "Where": [{"Condition": {"In": {"Expressions": [col_ref(a, prop)],
                                                   "Values": [[{"Literal": {"Value": f"'{v}'"}}] for v in values]}}}]},
        "type": "Categorical", "howCreated": 1, "isHiddenInViewMode": False}


def container(x, y, w, h, single_visual, filters=None, z=0):
    name = uuid.uuid4().hex[:20]
    cfg = {"name": name,
           "layouts": [{"id": 0, "position": {"x": x, "y": y, "z": z, "width": w, "height": h, "tabOrder": z}}],
           "singleVisual": single_visual}
    return {"x": x, "y": y, "z": z, "width": w, "height": h,
            "config": json.dumps(cfg), "filters": json.dumps(filters or [])}


def vc_title(text=None, measure=None, size=10, show=True):
    props = {"show": lit(show), "fontSize": lit(float(size)), "alignment": lit("center"),
             "fontColor": color("#252423")}
    if measure:
        props["text"] = measure_expr("Trend", measure)
    elif text is not None:
        props["text"] = lit(text)
    return {"title": [{"properties": props}]}


def textbox(x, y, w, h, runs, z=0, bg=None):
    """runs: list of paragraphs, each a list of (text, style dict)."""
    paragraphs = [{"textRuns": [{"value": t, "textStyle": s} for t, s in para]} for para in runs]
    sv = {"visualType": "textbox", "drillFilterOtherVisuals": True,
          "objects": {"general": [{"properties": {"paragraphs": paragraphs}}]}}
    if bg:
        sv["vcObjects"] = {"background": [{"properties": {"show": lit(True), "color": color(bg),
                                                          "transparency": lit(0.0)}}]}
    return container(x, y, w, h, sv, z=z)


def trend_chart(x, y, w, h, filters, title_measure="Period Title", z=0, show_legend=False, y_title=None):
    fields = [("col", "Trend", "DateTime"), ("meas", "Trend", "Actual Value"), ("meas", "Trend", "Reference Value")]
    sv = {
        "visualType": "lineChart",
        "projections": {"Category": [{"queryRef": "Trend.DateTime", "active": True}],
                        "Y": [{"queryRef": "Trend.Actual Value"}, {"queryRef": "Trend.Reference Value"}]},
        "prototypeQuery": query(fields, order_by=("Trend", "DateTime")),
        "drillFilterOtherVisuals": True,
        "objects": {
            # "Scalar" is what the Format pane calls "Continuous"
            "categoryAxis": [{"properties": {"axisType": lit("Scalar"), "showAxisTitle": lit(False),
                                             "fontSize": lit(7.0), "gridlineShow": lit(True)}}],
            "valueAxis": [{"properties": {"showAxisTitle": lit(y_title is not None), "fontSize": lit(7.0),
                                          **({"titleText": y_title, "titleFontSize": lit(8.0)} if y_title else {}),
                                          "gridlineShow": lit(True),
                                          "start": measure_expr("Trend", "Y Axis Min"),
                                          "end": measure_expr("Trend", "Y Axis Max")}}],
            "legend": [{"properties": {"show": lit(show_legend), "position": lit("Top"),
                                       "fontSize": lit(7.0)}}],
            "dataPoint": [
                {"properties": {"fill": color(ACTUAL_COLOR)}, "selector": {"metadata": "Trend.Actual Value"}},
                {"properties": {"fill": color(REFERENCE_COLOR)}, "selector": {"metadata": "Trend.Reference Value"}},
            ],
            "lineStyles": [
                {"properties": {"strokeWidth": lit(1), "showMarker": lit(False)}},
                {"properties": {"strokeWidth": lit(2)}, "selector": {"metadata": "Trend.Reference Value"}},
            ],
        },
        "vcObjects": vc_title(measure=title_measure, size=10),
    }
    return container(x, y, w, h, sv, filters=filters, z=z)


def card(x, y, w, h, measure, label, z=0, font=20):
    sv = {"visualType": "card",
          "projections": {"Values": [{"queryRef": f"Trend.{measure}"}]},
          "prototypeQuery": query([("meas", "Trend", measure)]),
          "drillFilterOtherVisuals": True,
          "objects": {"labels": [{"properties": {"fontSize": lit(float(font)), "color": color("#252423")}}],
                      "categoryLabels": [{"properties": {"show": lit(False)}}]},
          "vcObjects": {**vc_title(text=label, size=9),
                        "border": [{"properties": {"show": lit(True), "color": color("#D0D0D0")}}]}}
    return container(x, y, w, h, sv, z=z)


def title_card(x, y, w, h, measure, z=0):
    sv = {"visualType": "card",
          "projections": {"Values": [{"queryRef": f"Trend.{measure}"}]},
          "prototypeQuery": query([("meas", "Trend", measure)]),
          "drillFilterOtherVisuals": True,
          "objects": {"labels": [{"properties": {"fontSize": lit(13.0), "color": color("#252423")}}],
                      "categoryLabels": [{"properties": {"show": lit(False)}}]},
          "vcObjects": {"title": [{"properties": {"show": lit(False)}}]}}
    return container(x, y, w, h, sv, z=z)


def tag_slicer(x, y, w, h, default_tag_label, z=0):
    a = ALIAS["DimTag"]
    default_filter = {"Version": 2, "From": [{"Name": a, "Entity": "DimTag", "Type": 0}],
                      "Where": [{"Condition": {"In": {"Expressions": [col_ref(a, "TagLabel")],
                                                      "Values": [[{"Literal": {"Value": f"'{default_tag_label}'"}}]]}}}]}
    sv = {"visualType": "slicer",
          "projections": {"Values": [{"queryRef": "DimTag.TagLabel", "active": True}]},
          "prototypeQuery": query([("col", "DimTag", "TagLabel")]),
          "drillFilterOtherVisuals": True,
          "objects": {"data": [{"properties": {"mode": lit("Dropdown")}}],
                      "selection": [{"properties": {"strictSingleSelect": lit(True),
                                                    "selectAllCheckboxEnabled": lit(False)}}],
                      "general": [{"properties": {"filter": {"filter": default_filter}}}],
                      "header": [{"properties": {"show": lit(True), "text": lit("Variable"),
                                                 "fontSize": lit(10.0)}}]},
          "vcObjects": {"title": [{"properties": {"show": lit(False)}}]}}
    return container(x, y, w, h, sv, z=z)


def summary_table(x, y, w, h, z=0, by_tag=False):
    measures = ["Reference Value", "Average Value", "Min Value", "Max Value",
                "Deviation from Reference %", "% Time Above Reference"]
    rows = [("col", "DimTag", "ShortName")] if by_tag else []
    rows += [("col", "DimPeriod", "PeriodTitle")]
    fields = rows + [("meas", "Trend", m) for m in measures]
    sv = {"visualType": "tableEx",
          "projections": {"Values": [{"queryRef": f"{e}.{p}"} for _, e, p in fields]},
          "prototypeQuery": query(fields),
          "drillFilterOtherVisuals": True,
          "objects": {"grid": [{"properties": {"textSize": lit(9.0)}}],
                      "columnHeaders": [{"properties": {"fontSize": lit(9.0), "fontColor": color("#FFFFFF"),
                                                        "backColor": color("#1F4E79")}}]},
          "vcObjects": vc_title(text="Statistics per period (vs reference value)", size=10)}
    return container(x, y, w, h, sv, z=z)


def deviation_matrix(x, y, w, h, z=0):
    fields = [("col", "DimTag", "TagLabel"), ("col", "DimPeriod", "PeriodTitle"),
              ("meas", "Trend", "Average Value"), ("meas", "Trend", "Deviation from Reference %")]
    sv = {"visualType": "pivotTable",
          "projections": {"Rows": [{"queryRef": "DimTag.TagLabel", "active": True}],
                          "Columns": [{"queryRef": "DimPeriod.PeriodTitle", "active": True}],
                          "Values": [{"queryRef": "Trend.Average Value"},
                                     {"queryRef": "Trend.Deviation from Reference %"}]},
          "prototypeQuery": query(fields),
          "drillFilterOtherVisuals": True,
          "objects": {"grid": [{"properties": {"textSize": lit(9.0)}}],
                      "columnHeaders": [{"properties": {"fontSize": lit(9.0), "fontColor": color("#FFFFFF"),
                                                        "backColor": color("#1F4E79")}}],
                      "subTotals": [{"properties": {"rowSubtotals": lit(False), "columnSubtotals": lit(False)}}]},
          "vcObjects": vc_title(text="Average value and deviation from reference, per variable and period", size=11)}
    return container(x, y, w, h, sv, z=z)


H1 = {"fontSize": "18pt", "fontWeight": "bold", "color": "#1F4E79"}
SUB = {"fontSize": "10pt", "color": "#605E5C"}
ROWHDR = {"fontSize": "11pt", "fontWeight": "bold", "color": "#252423"}
ROWSUB = {"fontSize": "9pt", "color": "#605E5C"}
ROWUNIT = {"fontSize": "11pt", "fontWeight": "bold", "color": "#1F77B4"}


def section(name, display, ordinal, containers, width=1280, height=720, display_option=1):
    return {"name": name, "displayName": display, "displayOption": display_option,
            "width": width, "height": height, "ordinal": ordinal,
            "config": json.dumps({}), "filters": "[]", "visualContainers": containers}


def page_explorer():
    vc = []
    z = iter(range(0, 10000, 100))
    vc.append(textbox(16, 8, 820, 52, [[("MT Chiller 1 - Historical Trend vs Reference", H1)],
                                        [("Source: 2.Hist Value1 .. 2.Hist Value5 (1-min data, 2 days per period)", SUB)]],
                      z=next(z)))
    first_label = (f'{TAGS[0]["ShortName"]} - ' + TAGS[0]["Description"].replace("(CUP) Level 1 MT Chiller 1", "").strip()
                   + f' ({TAGS[0]["Unit"]})')
    vc.append(tag_slicer(900, 10, 364, 54, first_label, z=next(z)))
    vc.append(title_card(16, 66, 1248, 34, "Selected Tag Title", z=next(z)))
    cw, gap, cy = 120, 8, 106
    # KPI cards across the top
    kpis = [("Reference Value", "Reference Value"), ("Average Value", "Average (all periods)"),
            ("Min Value", "Min (all periods)"), ("Max Value", "Max (all periods)"),
            ("Deviation from Reference %", "Avg deviation vs reference"),
            ("% Time Above Reference", "% of time above reference")]
    kw = (1248 - gap * (len(kpis) - 1)) / len(kpis)
    for i, (m, lbl) in enumerate(kpis):
        vc.append(card(16 + i * (kw + gap), cy, kw, 70, m, lbl, z=next(z), font=18))
    # 5 period charts in one row, like the matplotlib subplots
    w = (1248 - gap * (N_PERIODS - 1)) / N_PERIODS
    for i, p in enumerate(PERIODS):
        vc.append(trend_chart(16 + i * (w + gap), 184, w, 330,
                              [in_filter("DimPeriod", "Period", [p])], z=next(z), show_legend=(i == 0),
                              y_title=measure_expr("Trend", "Unit Axis Title") if i == 0 else None))
    vc.append(summary_table(16, 522, 1248, 190, z=next(z)))
    return section("ReportSectionExplorer", "Trend Explorer", 1, vc)


def page_all_tags():
    vc = []
    z = iter(range(0, 100000, 100))
    row_h, hdr_h, gap, top = 270, 40, 8, 70
    vc.append(textbox(16, 8, 1248, 52, [[("MT Chiller 1 - All Variables, Hist_Value1 .. Hist_Value5", H1)],
                                        [("Blue = actual value (1-min), red = reference value from client. Y-axis is shared across the 5 periods "
                                          "of each variable. Units: kW cooling load, % power loading, kW/RT efficiency, "
                                          "m³/h flow, °C temperature.", SUB)]],
                      z=next(z)))
    w = (1248 - gap * (N_PERIODS - 1)) / N_PERIODS
    for r, t in enumerate(TAGS):
        y0 = top + r * (row_h + hdr_h + gap)
        ref = float(t["ReferenceValue"])
        ref_txt = f"{ref:,.3f}" if ref < 1 else f"{ref:,.2f}"
        vc.append(textbox(16, y0, 1248, hdr_h,
                          [[(f'Trend Graph - {t["Tag"]}', ROWHDR),
                            (f'  ({t["Unit"]})', ROWUNIT),
                            (f'    {t["Description"]}   |   Reference Value = {ref_txt} {t["Unit"]}', ROWSUB)]],
                          z=next(z), bg="#EEF3F8"))
        for i, p in enumerate(PERIODS):
            vc.append(trend_chart(16 + i * (w + gap), y0 + hdr_h, w, row_h,
                                  [in_filter("DimTag", "Tag", [t["Tag"]]), in_filter("DimPeriod", "Period", [p])],
                                  z=next(z), show_legend=(i == 0),
                                  y_title=lit(f'{t["ShortName"]} ({t["Unit"]})') if i == 0 else None))
    height = top + len(TAGS) * (row_h + hdr_h + gap) + 16
    return section("ReportSectionAllTags", "All Variables", 0, vc, height=height, display_option=2)


def page_summary():
    vc = [textbox(16, 8, 1248, 52, [[("Summary - Average vs Reference", H1)],
                                    [("Deviation % = (period average - reference) / reference", SUB)]]),
          deviation_matrix(16, 70, 1248, 330, z=100),
          summary_table(16, 410, 1248, 300, z=200, by_tag=True)]
    return section("ReportSectionSummary", "Summary", 2, vc)


def build_report(root):
    write_json(root / "definition.pbir", {
        "version": "4.0",
        "datasetReference": {"byPath": {"path": f"../{NAME}.SemanticModel"}}})
    write_json(root / ".platform", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "Report", "displayName": NAME},
        "config": {"version": "2.0", "logicalId": guid()}})
    # Power BI needs a base theme packaged with the report, otherwise the canvas stays blank
    theme_dir = root / "StaticResources" / "SharedResources" / "BaseThemes"
    theme_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy(HERE / "data" / "theme" / f"{BASE_THEME}.json", theme_dir / f"{BASE_THEME}.json")
    report = {
        "config": json.dumps({"version": "5.59",
                              "themeCollection": {"baseTheme": {"name": BASE_THEME, "version": "5.61", "type": 2}},
                              "activeSectionIndex": 0,
                              "defaultDrillFilterOtherVisuals": True,
                              "linguisticSchemaSyncVersion": 0,
                              "settings": {"useNewFilterPaneExperience": True,
                                           "allowChangeFilterTypes": True,
                                           "useStylableVisualContainerHeader": True,
                                           "exportDataMode": 1,
                                           "useEnhancedTooltips": True}}),
        "layoutOptimization": 0,
        "resourcePackages": [{"resourcePackage": {
            "disabled": False, "name": "SharedResources", "type": 2,
            "items": [{"name": BASE_THEME, "path": f"BaseThemes/{BASE_THEME}.json", "type": 202}]}}],
        "sections": [page_all_tags(), page_explorer(), page_summary()],
    }
    write_json(root / "report.json", report)


def main():
    for sub in (f"{NAME}.SemanticModel", f"{NAME}.Report"):
        shutil.rmtree(HERE / sub, ignore_errors=True)
    write_json(HERE / f"{NAME}.pbip", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
        "version": "1.0",
        "artifacts": [{"report": {"path": f"{NAME}.Report"}}],
        "settings": {"enableAutoRecovery": True}})
    build_model(HERE / f"{NAME}.SemanticModel")
    build_report(HERE / f"{NAME}.Report")
    write(HERE / ".gitignore", "**/.pbi/localSettings.json\n**/.pbi/cache.abf\n")
    print(f"Built {HERE / (NAME + '.pbip')}")


if __name__ == "__main__":
    main()

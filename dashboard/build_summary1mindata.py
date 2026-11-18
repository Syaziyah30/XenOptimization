"""Build the summary1mindata Power BI project (PBIP).

One page, 9 line charts: average of each MT Chiller 1 variable per period
(Hist_Value1-5) with the overall average as a dashed benchmark line.
Data: sheet 'Data 1min' of CL_Correlation_Hist1-5.xlsx (same source as the
Excel 'Summary' and 'Visualization' sheets).
"""
import json
import os
import shutil
import sys
import uuid

HERE = os.path.dirname(os.path.abspath(__file__))
NAME = "summary1mindata"
SOURCE_XLSX = os.path.join(HERE, "CL_Correlation_Hist1-5.xlsx")
REPORT = os.path.join(HERE, f"{NAME}.Report")
MODEL = os.path.join(HERE, f"{NAME}.SemanticModel")
PAGE = "ReportSectionSummary1min"
EXPLORER = "ReportSectionVariableExplorer"

VARIABLES = [  # order, variable (prefix of the Excel header), description, unit, format
    (1, "CL1.PV", "Cooling Load", "kW", "#,##0.00"),
    (2, "EF.PV2", "Efficiency", "kW/RT", "0.0000"),
    (3, "FIT1.PV", "Condenser Water Return Flow", "m³/h", "#,##0.00"),
    (4, "FIT2.PV", "Chilled Water Supply Flow", "m³/h", "#,##0.00"),
    (5, "TIT1.PV", "Condenser Water Supply Temp", "°C", "0.00"),
    (6, "TIT2.PV", "Condenser Water Return Temp", "°C", "0.00"),
    (7, "TIT3.PV", "Chilled Water Supply Temp", "°C", "0.000"),
    (8, "TIT4.PV", "Chilled Water Return Temp", "°C", "0.000"),
    (9, "EF.PV", "Efficiency (power loading)", "%", "0.00"),
]
PERIODS = [(1, "Hist_Value1", "Hist_Value1 (Feb 2024)"), (2, "Hist_Value2", "Hist_Value2 (Aug 2024)"),
           (3, "Hist_Value3", "Hist_Value3 (Feb 2025)"), (4, "Hist_Value4", "Hist_Value4 (Aug 2025)"),
           (5, "Hist_Value5", "Hist_Value5 (Feb 2026)")]


def gid():
    return str(uuid.uuid4())


def write(path, text):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)


def write_json(path, obj):
    write(path, json.dumps(obj, indent=2, ensure_ascii=False))


def lit(v):
    return {"expr": {"Literal": {"Value": v}}}


def m_text(s):
    return '"' + s.replace('"', '""') + '"'


# ---------------------------------------------------------------- semantic model
def build_model():
    d = os.path.join(MODEL, "definition")
    write_json(os.path.join(MODEL, ".platform"), {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "SemanticModel", "displayName": NAME},
        "config": {"version": "2.0", "logicalId": gid()}})
    write_json(os.path.join(MODEL, "definition.pbism"), {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
        "version": "4.2", "settings": {}})
    write(os.path.join(d, "database.tmdl"), "database\n\tcompatibilityLevel: 1606\n\n")
    write(os.path.join(d, "model.tmdl"),
          "model Model\n\tculture: en-US\n\tdefaultPowerBIDataSourceVersion: powerBI_V3\n\tsourceQueryCulture: en-US\n"
          "\tdataAccessOptions\n\t\tlegacyRedirects\n\t\treturnErrorValuesAsNull\n\n"
          "annotation __PBI_TimeIntelligenceEnabled = 0\n\n"
          'annotation PBI_QueryOrder = ["SourceFilePath","Data1min","DimVariable","DimPeriod"]\n\n'
          'annotation PBI_ProTooling = ["DevMode"]\n\n'
          "ref table Data1min\nref table DimVariable\nref table DimPeriod\n\n")
    write(os.path.join(d, "expressions.tmdl"),
          f'expression SourceFilePath = {m_text(SOURCE_XLSX)} meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]\n'
          f"\tlineageTag: {gid()}\n\n\tannotation PBI_ResultType = Text\n\n")
    write(os.path.join(d, "relationships.tmdl"),
          f"relationship {gid()}\n\tfromColumn: Data1min.Variable\n\ttoColumn: DimVariable.Variable\n\n"
          f"relationship {gid()}\n\tfromColumn: Data1min.Period\n\ttoColumn: DimPeriod.Period\n\n")

    fmt = 'SELECTEDVALUE(DimVariable[FormatString], "#,##0.00")'
    data_tmdl = f"""table Data1min
	lineageTag: {gid()}

	/// Average of the 1-minute readings in the current filter (one variable, one period)
	measure 'Period Average' = AVERAGE(Data1min[Value])
		lineageTag: {gid()}

		formatStringDefinition = {fmt}

	/// Average over all 5 periods (benchmark line)
	measure 'Overall Average' = CALCULATE(AVERAGE(Data1min[Value]), REMOVEFILTERS(DimPeriod))
		lineageTag: {gid()}

		formatStringDefinition = {fmt}

	measure 'Y Axis Min' =
			VAR lo = MINX(ALL(DimPeriod[Period]), [Period Average])
			VAR hi = MAXX(ALL(DimPeriod[Period]), [Period Average])
			RETURN lo - (hi - lo) * 0.35
		lineageTag: {gid()}

	measure 'Y Axis Max' =
			VAR lo = MINX(ALL(DimPeriod[Period]), [Period Average])
			VAR hi = MAXX(ALL(DimPeriod[Period]), [Period Average])
			RETURN hi + (hi - lo) * 0.35
		lineageTag: {gid()}
{{EXPLORER_MEASURES}}
	column Period
		dataType: string
		lineageTag: {gid()}
		summarizeBy: none
		sourceColumn: Period

	column DateTime
		dataType: dateTime
		formatString: General Date
		lineageTag: {gid()}
		summarizeBy: none
		sourceColumn: DateTime

	column Variable
		dataType: string
		lineageTag: {gid()}
		summarizeBy: none
		sourceColumn: Variable

	column Value
		dataType: double
		lineageTag: {gid()}
		summarizeBy: none
		sourceColumn: Value

	partition Data1min = m
		mode: import
		source =
				let
				    Source = Excel.Workbook(File.Contents(SourceFilePath), null, true),
				    Sheet = Source{{[Item = "Data 1min", Kind = "Sheet"]}}[Data],
				    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),
				    Unpivoted = Table.UnpivotOtherColumns(Promoted, {{"Period", "DateTime"}}, "Header", "Value"),
				    WithVariable = Table.AddColumn(Unpivoted, "Variable", each Text.BeforeDelimiter([Header], " "), type text),
				    Typed = Table.TransformColumnTypes(WithVariable, {{{{"Period", type text}}, {{"DateTime", type datetime}}, {{"Value", type number}}}}, "en-US"),
				    Result = Table.SelectColumns(Typed, {{"Period", "DateTime", "Variable", "Value"}})
				in
				    Result

	annotation PBI_ResultType = Table

"""
    write(os.path.join(d, "tables", "Data1min.tmdl"), data_tmdl.replace("{EXPLORER_MEASURES}", explorer_measures()))

    rows = ", ".join("{" + f'{o}, {m_text(v)}, {m_text(ds)}, {m_text(u)}, {m_text(f"{v} – {ds} ({u})")}, {m_text(f)}' + "}"
                     for o, v, ds, u, f in VARIABLES)
    var_tmdl = f"""table DimVariable
	lineageTag: {gid()}

	column Order
		dataType: int64
		isHidden
		formatString: 0
		lineageTag: {gid()}
		summarizeBy: none
		sourceColumn: Order

	column Variable
		dataType: string
		lineageTag: {gid()}
		summarizeBy: none
		sourceColumn: Variable
		sortByColumn: Order

	column Description
		dataType: string
		lineageTag: {gid()}
		summarizeBy: none
		sourceColumn: Description
		sortByColumn: Order

	column Unit
		dataType: string
		lineageTag: {gid()}
		summarizeBy: none
		sourceColumn: Unit

	column ChartTitle
		dataType: string
		lineageTag: {gid()}
		summarizeBy: none
		sourceColumn: ChartTitle
		sortByColumn: Order

	column FormatString
		dataType: string
		isHidden
		lineageTag: {gid()}
		summarizeBy: none
		sourceColumn: FormatString

	partition DimVariable = m
		mode: import
		source =
				let
				    Source = #table(
				        type table [Order = Int64.Type, Variable = text, Description = text, Unit = text, ChartTitle = text, FormatString = text],
				        {{{rows}}})
				in
				    Source

	annotation PBI_ResultType = Table

"""
    write(os.path.join(d, "tables", "DimVariable.tmdl"), var_tmdl)

    prow = ", ".join("{" + f"{o}, {m_text(p)}, {m_text(lab)}" + "}" for o, p, lab in PERIODS)
    per_tmdl = f"""table DimPeriod
	lineageTag: {gid()}

	column Order
		dataType: int64
		isHidden
		formatString: 0
		lineageTag: {gid()}
		summarizeBy: none
		sourceColumn: Order

	column Period
		dataType: string
		lineageTag: {gid()}
		summarizeBy: none
		sourceColumn: Period
		sortByColumn: Order

	/// Period with its month, e.g. "Hist_Value1 (Feb 2024)"
	column PeriodLabel
		dataType: string
		lineageTag: {gid()}
		summarizeBy: none
		sourceColumn: PeriodLabel
		sortByColumn: Order

	partition DimPeriod = m
		mode: import
		source =
				let
				    Source = #table(type table [Order = Int64.Type, Period = text, PeriodLabel = text], {{{prow}}})
				in
				    Source

	annotation PBI_ResultType = Table

"""
    write(os.path.join(d, "tables", "DimPeriod.tmdl"), per_tmdl)


def explorer_measures():
    """Measures used by the Variable Explorer page (title, difference vs benchmark)."""
    return f"""
	/// Chart title for the variable ticked in the slicer
	measure 'Selected Chart Title' = SELECTEDVALUE(DimVariable[ChartTitle], "Tick one variable on the left")
		lineageTag: {gid()}

	/// Period average minus overall average, as % of the overall average
	measure 'Diff vs Overall %' = DIVIDE([Period Average] - [Overall Average], [Overall Average])
		lineageTag: {gid()}
		formatString: +0.00%;-0.00%;0.00%
"""


# ---------------------------------------------------------------- report
def measure(prop):
    return {"Measure": {"Expression": {"SourceRef": {"Entity": "Data1min"}}, "Property": prop}}


def column(entity, prop):
    return {"Column": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}}


def textbox(name, x, y, w, h, text, size, color, bold=False):
    style = {"fontSize": f"{size}pt", "color": color}
    if bold:
        style["fontWeight"] = "bold"
    return {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.14.0/schema.json",
        "name": name,
        "position": {"x": x, "y": y, "z": 0, "height": h, "width": w, "tabOrder": 0},
        "visual": {"visualType": "textbox",
                   "objects": {"general": [{"properties": {"paragraphs": [{"textRuns": [{"value": text, "textStyle": style}]}]}}]},
                   "drillFilterOtherVisuals": True}}


def line_chart(name, idx, x, y, w, h, var, title, label_size="8.0D", axis_size="8.0D"):
    """Period average + dashed overall average. var=None: no visual filter (a slicer
    decides) and title=None: the title comes from the 'Selected Chart Title' measure."""
    title_expr = {"expr": measure("Selected Chart Title")} if title is None else lit("'" + title.replace("'", "''") + "'")
    vis = {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.14.0/schema.json",
        "name": name,
        "position": {"x": x, "y": y, "z": 1000 + idx, "height": h, "width": w, "tabOrder": 1000 + idx},
        "visual": {
            "visualType": "lineChart",
            "query": {
                "queryState": {
                    "Category": {"projections": [{"field": column("DimPeriod", "PeriodLabel"), "queryRef": "DimPeriod.PeriodLabel",
                                                  "nativeQueryRef": "PeriodLabel", "active": True}]},
                    "Y": {"projections": [
                        {"field": measure("Period Average"), "queryRef": "Data1min.Period Average", "nativeQueryRef": "Period Average"},
                        {"field": measure("Overall Average"), "queryRef": "Data1min.Overall Average", "nativeQueryRef": "Overall Average"}]}},
                "sortDefinition": {"sort": [{"field": column("DimPeriod", "PeriodLabel"), "direction": "Ascending"}]}},
            "objects": {
                "categoryAxis": [{"properties": {"showAxisTitle": lit("false"), "fontSize": lit(axis_size), "concatenateLabels": lit("false")}}],
                "valueAxis": [{"properties": {"showAxisTitle": lit("false"), "fontSize": lit(axis_size), "gridlineShow": lit("true"),
                                              "start": {"expr": measure("Y Axis Min")}, "end": {"expr": measure("Y Axis Max")}}}],
                "legend": [{"properties": {"show": lit("true"), "position": lit("'Bottom'"), "fontSize": lit("8.0D")}}],
                "dataPoint": [
                    {"properties": {"fill": {"solid": {"color": lit("'#1F6FB2'")}}}, "selector": {"metadata": "Data1min.Period Average"}},
                    {"properties": {"fill": {"solid": {"color": lit("'#ED7D31'")}}}, "selector": {"metadata": "Data1min.Overall Average"}}],
                "lineStyles": [
                    {"properties": {"strokeWidth": lit("3L"), "showMarker": lit("true"), "lineChartType": lit("'linear'")}},
                    {"properties": {"strokeWidth": lit("2L"), "lineStyle": lit("'dashed'"), "showMarker": lit("false")},
                     "selector": {"metadata": "Data1min.Overall Average"}}],
                "labels": [
                    {"properties": {"show": lit("true"), "fontSize": lit(label_size), "labelPosition": lit("'Above'")}},
                    {"properties": {"showSeries": lit("false")}, "selector": {"metadata": "Data1min.Overall Average"}}]},
            "visualContainerObjects": {
                "title": [{"properties": {"show": lit("true"), "text": title_expr,
                                          "fontSize": lit("11.0D"), "bold": lit("true"), "alignment": lit("'center'"),
                                          "fontColor": {"solid": {"color": lit("'#14213D'")}}}}],
                "subTitle": [{"properties": {"show": lit("true"), "fontSize": lit("8.0D"), "alignment": lit("'center'"),
                                             "text": lit("'Dashed orange = overall average (Hist_Value1–5)'")}}],
                "border": [{"properties": {"show": lit("true"), "color": {"solid": {"color": lit("'#D0D7E2'")}}}}],
                "background": [{"properties": {"show": lit("true"), "color": {"solid": {"color": lit("'#FFFFFF'")}}, "transparency": lit("0D")}}]},
            "drillFilterOtherVisuals": True}}
    if var is not None:
        vis["filterConfig"] = {"filters": [{
            "name": f"varfilter{idx:02d}",
            "field": column("DimVariable", "Variable"),
            "type": "Categorical",
            "filter": in_filter("DimVariable", "Variable", var),
            "howCreated": "User"}]}
    return vis


def in_filter(entity, prop, value):
    return {"Version": 2, "From": [{"Name": "v", "Entity": entity, "Type": 0}],
            "Where": [{"Condition": {"In": {
                "Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "v"}}, "Property": prop}}],
                "Values": [[{"Literal": {"Value": "'" + value.replace("'", "''") + "'"}}]]}}}]}


def variable_slicer(name, x, y, w, h):
    first = VARIABLES[0]
    default = f"{first[1]} – {first[2]} ({first[3]})"
    return {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.14.0/schema.json",
        "name": name,
        "position": {"x": x, "y": y, "z": 200, "height": h, "width": w, "tabOrder": 200},
        "visual": {
            "visualType": "slicer",
            "query": {"queryState": {"Values": {"projections": [{
                "field": column("DimVariable", "ChartTitle"), "queryRef": "DimVariable.ChartTitle",
                "nativeQueryRef": "ChartTitle", "active": True}]}}},
            "objects": {
                "data": [{"properties": {"mode": lit("'Basic'")}}],
                "selection": [{"properties": {"strictSingleSelect": lit("false"), "singleSelect": lit("true"),
                                              "selectAllCheckboxEnabled": lit("false")}}],
                "header": [{"properties": {"show": lit("true"), "text": lit("'Tick a variable'"),
                                           "fontColor": {"solid": {"color": lit("'#1F4E79'")}}, "textSize": lit("12D"), "bold": lit("true")}}],
                "items": [{"properties": {"textSize": lit("11D"), "fontColor": {"solid": {"color": lit("'#252423'")}}}}],
                "general": [{"properties": {"filter": {"filter": in_filter("DimVariable", "ChartTitle", default)}}}]},
            "visualContainerObjects": {
                "title": [{"properties": {"show": lit("false")}}],
                "background": [{"properties": {"show": lit("true"), "color": {"solid": {"color": lit("'#FFFFFF'")}}}}],
                "border": [{"properties": {"show": lit("true"), "color": {"solid": {"color": lit("'#D0D7E2'")}}, "radius": lit("6D")}}]},
            "drillFilterOtherVisuals": True}}


def period_table(name, x, y, w, h):
    proj = [
        {"field": column("DimPeriod", "Period"), "queryRef": "DimPeriod.Period", "nativeQueryRef": "Period", "displayName": "Period"},
        {"field": measure("Period Average"), "queryRef": "Data1min.Period Average", "nativeQueryRef": "Period Average", "displayName": "Average"},
        {"field": measure("Diff vs Overall %"), "queryRef": "Data1min.Diff vs Overall %", "nativeQueryRef": "Diff vs Overall %",
         "displayName": "vs overall"}]
    return {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.14.0/schema.json",
        "name": name,
        "position": {"x": x, "y": y, "z": 300, "height": h, "width": w, "tabOrder": 300},
        "visual": {
            "visualType": "tableEx",
            "query": {"queryState": {"Values": {"projections": proj}},
                      "sortDefinition": {"sort": [{"field": column("DimPeriod", "Period"), "direction": "Ascending"}]}},
            "objects": {
                "columnWidth": [
                    {"properties": {"value": lit("120D")}, "selector": {"metadata": "DimPeriod.Period"}},
                    {"properties": {"value": lit("95D")}, "selector": {"metadata": "Data1min.Period Average"}},
                    {"properties": {"value": lit("85D")}, "selector": {"metadata": "Data1min.Diff vs Overall %"}}],
                "columnHeaders": [{"properties": {"fontSize": lit("10D"), "fontColor": {"solid": {"color": lit("'#FFFFFF'")}},
                                                  "backColor": {"solid": {"color": lit("'#14213D'")}}}}],
                "values": [{"properties": {"fontSize": lit("10D")}}],
                "total": [{"properties": {"totals": lit("false")}}]},
            "visualContainerObjects": {
                "title": [{"properties": {"show": lit("true"), "text": lit("'Values for the ticked variable'"), "fontSize": lit("11.0D"),
                                          "bold": lit("true"), "fontColor": {"solid": {"color": lit("'#1F4E79'")}}}}],
                "background": [{"properties": {"show": lit("true"), "color": {"solid": {"color": lit("'#FFFFFF'")}}}}],
                "border": [{"properties": {"show": lit("true"), "color": {"solid": {"color": lit("'#D0D7E2'")}}, "radius": lit("6D")}}]},
            "drillFilterOtherVisuals": True}}


def build_explorer_page():
    """Second page: one big chart, the variable is chosen in a slicer on the left."""
    d = os.path.join(REPORT, "definition", "pages", EXPLORER)
    if os.path.exists(d):
        shutil.rmtree(d)
    PW, PH, X0, Y0, LW, GAP = 1290, 760, 15, 80, 330, 15
    write_json(os.path.join(d, "page.json"), {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json",
        "name": EXPLORER, "displayName": "Variable Explorer", "displayOption": "FitToPage", "height": PH, "width": PW,
        "objects": {"background": [{"properties": {"color": {"solid": {"color": lit("'#EEF2F7'")}}, "transparency": lit("0D")}}]}})
    v = os.path.join(d, "visuals")
    write_json(os.path.join(v, "vextitle00000001", "visual.json"),
               textbox("vextitle00000001", X0, 8, PW - 2 * X0, 36,
                       "MT Chiller 1 – Variable Explorer: average by period vs overall average", 16, "#14213D", True))
    write_json(os.path.join(v, "vexsubtitle00001", "visual.json"),
               textbox("vexsubtitle00001", X0, 44, PW - 2 * X0, 28,
                       "Tick a variable on the left. Blue = average of the 1-minute readings in each period. Dashed orange = overall average of all 5 periods.",
                       9, "#605E5C"))
    Y0 = 86  # a little more room under the subtitle
    slicer_h = 330
    write_json(os.path.join(v, "vexslicer0000001", "visual.json"), variable_slicer("vexslicer0000001", X0, Y0, LW, slicer_h))
    write_json(os.path.join(v, "vextable00000001", "visual.json"),
               period_table("vextable00000001", X0, Y0 + slicer_h + GAP, LW, PH - Y0 - slicer_h - GAP - X0))
    write_json(os.path.join(v, "vexchart00000001", "visual.json"),
               line_chart("vexchart00000001", 1, X0 + LW + GAP, Y0, PW - 2 * X0 - LW - GAP, PH - Y0 - X0, None, None,
                          label_size="11.0D", axis_size="11.0D"))


def add_explorer_to_existing():
    """Add only the Variable Explorer page and its measures, keeping the files Power BI saved."""
    tm = os.path.join(MODEL, "definition", "tables", "Data1min.tmdl")
    with open(tm, encoding="utf-8") as f:
        text = f.read()
    if "Selected Chart Title" not in text:
        anchor = "\n\tcolumn Period\n"
        text = text.replace(anchor, explorer_measures() + anchor, 1)
        write(tm, text)
    build_explorer_page()
    pj = os.path.join(REPORT, "definition", "pages", "pages.json")
    with open(pj, encoding="utf-8") as f:
        pages = json.load(f)
    if EXPLORER not in pages["pageOrder"]:
        pages["pageOrder"].append(EXPLORER)
    pages["activePageName"] = EXPLORER
    write_json(pj, pages)
    print("added", EXPLORER)


def build_report():
    d = os.path.join(REPORT, "definition")
    write_json(os.path.join(REPORT, ".platform"), {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "Report", "displayName": NAME},
        "config": {"version": "2.0", "logicalId": gid()}})
    write_json(os.path.join(REPORT, "definition.pbir"), {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
        "version": "4.0", "datasetReference": {"byPath": {"path": f"../{NAME}.SemanticModel"}}})
    write_json(os.path.join(d, "version.json"), {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json",
        "version": "2.0.0"})
    # same theme and report settings as ChillerTrend
    with open(os.path.join(HERE, "CL_Dashboard.Report", "definition", "report.json"), encoding="utf-8") as f:
        write(os.path.join(d, "report.json"), f.read())
    theme_src = os.path.join(HERE, "CL_Dashboard.Report", "StaticResources")
    shutil.copytree(theme_src, os.path.join(REPORT, "StaticResources"), dirs_exist_ok=True)

    write_json(os.path.join(d, "pages", "pages.json"), {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.1.0/schema.json",
        "pageOrder": [PAGE], "activePageName": PAGE})
    W, H, GAP, X0, Y0 = 410, 290, 15, 15, 80
    write_json(os.path.join(d, "pages", PAGE, "page.json"), {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json",
        "name": PAGE, "displayName": "Summary 1-min Averages", "displayOption": "FitToWidth",
        "height": Y0 + 3 * H + 3 * GAP, "width": 2 * X0 + 3 * W + 2 * GAP,
        "objects": {"background": [{"properties": {"color": {"solid": {"color": lit("'#EEF2F7'")}}, "transparency": lit("0D")}}]}})
    vdir = os.path.join(d, "pages", PAGE, "visuals")
    width = 2 * X0 + 3 * W + 2 * GAP
    write_json(os.path.join(vdir, "s1mtitle00000001", "visual.json"),
               textbox("s1mtitle00000001", X0, 8, width - 2 * X0, 36,
                       "MT Chiller 1 – Average by period vs overall average (1-minute client data)", 16, "#14213D", True))
    write_json(os.path.join(vdir, "s1msubtitle00001", "visual.json"),
               textbox("s1msubtitle00001", X0, 44, width - 2 * X0, 28,
                       "Blue = average of the 1-minute readings in each period (2,880 per period). Dashed orange = overall average of all 5 periods. "
                       "Source: 'Data 1min' sheet of CL_Correlation_Hist1-5.xlsx – the same numbers as the Excel 'Summary' sheet.",
                       9, "#605E5C"))
    for i, (o, var, ds, unit, _) in enumerate(VARIABLES):
        name = f"s1mchart{i+1:08d}"
        x = X0 + (i % 3) * (W + GAP)
        y = Y0 + (i // 3) * (H + GAP)
        write_json(os.path.join(vdir, name, "visual.json"), line_chart(name, i + 1, x, y, W, H, var, f"{var} – {ds} ({unit})"))


def main():
    if "--add-explorer" in sys.argv:
        add_explorer_to_existing()
        return
    for p in (REPORT, MODEL):
        if os.path.exists(p):
            shutil.rmtree(p)
    build_model()
    build_report()
    build_explorer_page()
    pj = os.path.join(REPORT, "definition", "pages", "pages.json")
    write_json(pj, {"$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.1.0/schema.json",
                    "pageOrder": [PAGE, EXPLORER], "activePageName": PAGE})
    write_json(os.path.join(HERE, f"{NAME}.pbip"), {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
        "version": "1.0", "artifacts": [{"report": {"path": f"{NAME}.Report"}}], "settings": {"enableAutoRecovery": True}})
    print("built", os.path.join(HERE, f"{NAME}.pbip"))


if __name__ == "__main__":
    main()

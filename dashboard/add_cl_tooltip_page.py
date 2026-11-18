"""Custom tooltip for the CL Combined Dashboard chart in CL_Dashboard.

Power BI's default tooltip left-aligns values, so this adds a hidden tooltip page with a
two-column table (name | value) where values are right-aligned with 3 decimals, and points
the chart's tooltip at it. Also puts the 19 chart measures back to a plain #,0.000 format.
Run with CL_Dashboard closed. Safe to run again.
"""
import json
import os
import re
import shutil
import uuid

HERE = os.path.dirname(os.path.abspath(__file__))
MODEL = os.path.join(HERE, "CL_Dashboard.SemanticModel", "definition")
PAGES = os.path.join(HERE, "CL_Dashboard.Report", "definition", "pages")
CHART = os.path.join(PAGES, "c6e08ab2c3a08490c2eb", "visuals", "e665990ba32089013a78", "visual.json")
TIP_PAGE = "ReportSectionCLTooltip"
ROWS_TABLE = "CL Tooltip Rows"

# (short name, unit, value measure, reference measure) in the order of the chart
VARS = [("CL1.PV", "kW", "CL1.PV (kW)", "CL1.PV Avg Reference (kW)")] + [
    (s, u, f"CLc {s}", f"CLc {s} Reference")
    for s, u in [("EF.PV2", "kW/RT"), ("FIT1.PV", "m³/h"), ("FIT2.PV", "m³/h"), ("TIT1.PV", "°C"), ("TIT2.PV", "°C"),
                 ("TIT3.PV", "°C"), ("TIT4.PV", "°C"), ("EF.PV", "%")]]
CHART_MEASURES = [m for v in VARS for m in v[2:]] + ["CL1.PV Avg Reference label"]


def gid():
    return str(uuid.uuid4())


def lit(v):
    return {"expr": {"Literal": {"Value": v}}}


def col(entity, prop):
    return {"Column": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}}


def meas(prop):
    return {"Measure": {"Expression": {"SourceRef": {"Entity": "Trend"}}, "Property": prop}}


def write_json(path, obj):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        json.dump(obj, f, indent=2, ensure_ascii=False)


def m_text(s):
    return '"' + s.replace('"', '""') + '"'


# ---------------------------------------------------------------- model
def rows():
    out, order = [], 1
    for s, u, _, _ in VARS:
        ref_label = f"{s} reference (kW)" if s == "CL1.PV" else f"{s} reference"
        out += [(order, f"{s} ({u})", s, "value"), (order + 1, ref_label, s, "ref")]
        order += 2
    return out


def update_model():
    tables = os.path.join(MODEL, "tables")
    data = ", ".join("{" + f"{o}, {m_text(lab)}, {m_text(s)}, {m_text(k)}" + "}" for o, lab, s, k in rows())
    with open(os.path.join(tables, f"{ROWS_TABLE}.tmdl"), "w", encoding="utf-8", newline="\n") as f:
        f.write(f"""/// Row labels for the custom tooltip of the CL Combined Dashboard chart
table '{ROWS_TABLE}'
	isHidden
	lineageTag: {gid()}

	column Order
		dataType: int64
		isHidden
		formatString: 0
		lineageTag: {gid()}
		summarizeBy: none
		sourceColumn: Order

	column Label
		dataType: string
		lineageTag: {gid()}
		summarizeBy: none
		sourceColumn: Label
		sortByColumn: Order

	column Short
		dataType: string
		isHidden
		lineageTag: {gid()}
		summarizeBy: none
		sourceColumn: Short

	column Kind
		dataType: string
		isHidden
		lineageTag: {gid()}
		summarizeBy: none
		sourceColumn: Kind

	partition '{ROWS_TABLE}' = m
		mode: import
		source =
				let
				    Source = #table(type table [Order = Int64.Type, Label = text, Short = text, Kind = text], {{{data}}})
				in
				    Source

	annotation PBI_ResultType = Table

""")
    mpath = os.path.join(MODEL, "model.tmdl")
    with open(mpath, encoding="utf-8") as f:
        model = f.read()
    if f"ref table '{ROWS_TABLE}'" not in model:
        model = re.sub(r"(ref table [^\n]+\n)(?!ref table)", rf"\1ref table '{ROWS_TABLE}'\n", model, count=1)
        with open(mpath, "w", encoding="utf-8", newline="\n") as f:
            f.write(model)

    tpath = os.path.join(tables, "Trend.tmdl")
    with open(tpath, encoding="utf-8") as f:
        t = f.read()
    # 1) plain 3-decimal format on the chart measures (the space padding did not survive in tooltips)
    for n in CHART_MEASURES:
        pat = re.compile(r"(\n\tmeasure '" + re.escape(n) + r"' =.*?)(?=\n\t(?:measure|column|partition|///|annotation)|\Z)", re.S)
        m = pat.search(t)
        body = re.sub(r"\n\t\tformatString(Definition = |: )[^\n]*\n?", "\n", m.group(1)).rstrip("\n")
        body = re.sub(r"(\n\t\tlineageTag: [^\n]*)", r"\1\n\t\tformatString: #,0.000", body, count=1) + "\n"
        t = t[:m.start(1)] + body + t[m.end(1):]
    # 2) tooltip measures
    if "'CL Tooltip Value'" not in t:
        cases = ",\n".join(f'\t\t\t    "{s}|value", [{v}],\n\t\t\t    "{s}|ref", [{r}]' for s, _, v, r in VARS)
        add = f"""
	/// Value for one row of the custom CL tooltip (blank when that variable is not ticked)
	measure 'CL Tooltip Value' =
			VAR k = SELECTEDVALUE ( '{ROWS_TABLE}'[Short] ) & "|" & SELECTEDVALUE ( '{ROWS_TABLE}'[Kind] )
			RETURN
			SWITCH (
			    k,
{cases}
			)
		formatString: #,0.000
		lineageTag: {gid()}

	/// Title of the custom CL tooltip: the hovered 30-minute block
	measure 'CL Tooltip Title' = FORMAT ( MIN ( 'Trend'[DateTime 30min] ), "d/m/yyyy  hh:mm" )
		lineageTag: {gid()}
"""
        anchor = "\n\tcolumn Period\n"
        t = t.replace(anchor, add + anchor, 1)
    with open(tpath, "w", encoding="utf-8", newline="\n") as f:
        f.write(t)


# ---------------------------------------------------------------- report
def tooltip_table(name, w, h):
    proj = [{"field": col(ROWS_TABLE, "Label"), "queryRef": f"{ROWS_TABLE}.Label", "nativeQueryRef": "Label", "displayName": " "},
            {"field": meas("CL Tooltip Value"), "queryRef": "Trend.CL Tooltip Value", "nativeQueryRef": "CL Tooltip Value",
             "displayName": "  "}]
    return {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.14.0/schema.json",
        "name": name,
        "position": {"x": 0, "y": 0, "z": 0, "height": h, "width": w, "tabOrder": 0},
        "visual": {
            "visualType": "tableEx",
            "query": {"queryState": {"Values": {"projections": proj}},
                      "sortDefinition": {"sort": [{"field": col(ROWS_TABLE, "Label"), "direction": "Ascending"}]}},
            "objects": {
                "grid": [{"properties": {"gridVertical": lit("false"), "gridHorizontal": lit("false"),
                                         "outlineColor": {"solid": {"color": lit("'#FFFFFF'")}}, "rowPadding": lit("3D")}}],
                "columnHeaders": [{"properties": {"fontSize": lit("6D"), "fontColor": {"solid": {"color": lit("'#FFFFFF'")}},
                                                  "backColor": {"solid": {"color": lit("'#FFFFFF'")}}, "outline": lit("'None'")}}],
                "values": [{"properties": {"fontSize": lit("10D"), "fontColor": {"solid": {"color": lit("'#252423'")}},
                                           "backColorPrimary": {"solid": {"color": lit("'#FFFFFF'")}},
                                           "backColorSecondary": {"solid": {"color": lit("'#FFFFFF'")}}, "outline": lit("'None'")}}],
                "columnWidth": [{"properties": {"value": lit("170D")}, "selector": {"metadata": f"{ROWS_TABLE}.Label"}},
                                {"properties": {"value": lit("95D")}, "selector": {"metadata": "Trend.CL Tooltip Value"}}],
                "columnFormatting": [
                    {"properties": {"alignment": lit("'Right'"), "fontColor": {"solid": {"color": lit("'#252423'")}}},
                     "selector": {"metadata": f"{ROWS_TABLE}.Label"}},
                    {"properties": {"alignment": lit("'Right'")}, "selector": {"metadata": "Trend.CL Tooltip Value"}}],
                "total": [{"properties": {"totals": lit("false")}}]},
            "visualContainerObjects": {
                "title": [{"properties": {"show": lit("true"), "text": {"expr": meas("CL Tooltip Title")}, "fontSize": lit("10D"),
                                          "bold": lit("true"), "fontColor": {"solid": {"color": lit("'#252423'")}}}}],
                "background": [{"properties": {"show": lit("true"), "color": {"solid": {"color": lit("'#FFFFFF'")}}}}],
                "border": [{"properties": {"show": lit("false")}}]},
            "drillFilterOtherVisuals": True}}


def update_report():
    W, H = 300, 230
    d = os.path.join(PAGES, TIP_PAGE)
    if os.path.exists(d):
        shutil.rmtree(d)
    write_json(os.path.join(d, "page.json"), {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json",
        "name": TIP_PAGE, "displayName": "CL Tooltip", "displayOption": "ActualSize", "height": H, "width": W,
        "type": "Tooltip", "visibility": "HiddenInViewMode",
        "pageBinding": {"name": "CLTooltipBinding", "type": "Tooltip", "parameters": []},
        "objects": {"background": [{"properties": {"color": {"solid": {"color": lit("'#FFFFFF'")}}, "transparency": lit("0D")}}]}})
    write_json(os.path.join(d, "visuals", "cltipTable0000001", "visual.json"), tooltip_table("cltipTable0000001", W, H))

    pj = os.path.join(PAGES, "pages.json")
    with open(pj, encoding="utf-8") as f:
        pages = json.load(f)
    if TIP_PAGE not in pages["pageOrder"]:
        pages["pageOrder"].append(TIP_PAGE)
    write_json(pj, pages)

    with open(CHART, encoding="utf-8") as f:
        chart = json.load(f)
    vco = chart["visual"].setdefault("visualContainerObjects", {})
    vco["visualTooltip"] = [{"properties": {"show": lit("true"), "type": lit("'ReportPage'"), "section": lit(f"'{TIP_PAGE}'")}}]
    write_json(CHART, chart)


if __name__ == "__main__":
    update_model()
    update_report()
    print("done")

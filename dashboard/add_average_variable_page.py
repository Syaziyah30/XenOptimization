"""Copy the 'Variable Explorer' page of summary1mindata into CL_Dashboard as 'Average Variable'.

The CL_Dashboard model already holds the same 1-minute data (table Trend), so the copied
visuals are re-pointed to Trend / DimTag / DimPeriod and to new 'AV ...' measures that are
added to Trend. Existing pages and measures are not touched. Safe to run again.
"""
import json
import os
import shutil
import uuid

HERE = os.path.dirname(os.path.abspath(__file__))
SRC_PAGE = os.path.join(HERE, "summary1mindata.Report", "definition", "pages", "ReportSectionVariableExplorer")
CL_PAGES = os.path.join(HERE, "CL_Dashboard.Report", "definition", "pages")
NEW_PAGE = "ReportSectionAverageVariable"
TREND_TMDL = os.path.join(HERE, "CL_Dashboard.SemanticModel", "definition", "tables", "Trend.tmdl")

MEASURES = {  # summary1mindata measure -> CL_Dashboard measure
    "Period Average": "AV Period Average",
    "Overall Average": "AV Overall Average",
    "Y Axis Min": "AV Y Axis Min",
    "Y Axis Max": "AV Y Axis Max",
    "Selected Chart Title": "AV Chart Title",
    "Diff vs Overall %": "AV Diff vs Overall %",
}
EXACT = {"Data1min": "Trend", "DimVariable": "DimTag", "ChartTitle": "TagLabel",
         "DimVariable.ChartTitle": "DimTag.TagLabel",
         "'CL1.PV – Cooling Load (kW)'": "'CL1.PV - Cooling Load (kW)'",
         "MT Chiller 1 – Variable Explorer: average by period vs overall average":
             "MT Chiller 1 – Average Variable: average by period vs overall average"}
EXACT.update(MEASURES)
EXACT.update({f"Data1min.{k}": f"Trend.{v}" for k, v in MEASURES.items()})

FMT = ('SWITCH ( SELECTEDVALUE ( DimTag[ShortName] ), "EF.PV2", "0.0000", "TIT3.PV", "0.000", "TIT4.PV", "0.000", '
       '"TIT1.PV", "0.00", "TIT2.PV", "0.00", "EF.PV", "0.00", "#,0.00" )')
ALL_PERIODS = "REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( 'Trend'[Period] )"


def gid():
    return str(uuid.uuid4())


def measures_tmdl():
    return f"""
	/// Average Variable page: average of the 1-minute readings for the ticked variable and period
	measure 'AV Period Average' = AVERAGE ( 'Trend'[Value] )
		lineageTag: {gid()}

		formatStringDefinition = {FMT}

	/// Average Variable page: average over all 5 periods (benchmark line)
	measure 'AV Overall Average' = CALCULATE ( AVERAGE ( 'Trend'[Value] ), {ALL_PERIODS} )
		lineageTag: {gid()}

		formatStringDefinition = {FMT}

	measure 'AV Y Axis Min' =
			VAR lo = MINX ( ALL ( DimPeriod[Period] ), [AV Period Average] )
			VAR hi = MAXX ( ALL ( DimPeriod[Period] ), [AV Period Average] )
			RETURN lo - ( hi - lo ) * 0.35
		lineageTag: {gid()}

	measure 'AV Y Axis Max' =
			VAR lo = MINX ( ALL ( DimPeriod[Period] ), [AV Period Average] )
			VAR hi = MAXX ( ALL ( DimPeriod[Period] ), [AV Period Average] )
			RETURN hi + ( hi - lo ) * 0.35
		lineageTag: {gid()}

	measure 'AV Chart Title' = SELECTEDVALUE ( DimTag[TagLabel], "Tick one variable on the left" )
		lineageTag: {gid()}

	measure 'AV Diff vs Overall %' = DIVIDE ( [AV Period Average] - [AV Overall Average], [AV Overall Average] )
		formatString: +0.00%;-0.00%;0.00%
		lineageTag: {gid()}
"""


def remap(node):
    if isinstance(node, dict):
        return {k: remap(v) for k, v in node.items()}
    if isinstance(node, list):
        return [remap(v) for v in node]
    if isinstance(node, str):
        return EXACT.get(node, node)
    return node


def main():
    with open(TREND_TMDL, encoding="utf-8") as f:
        text = f.read()
    if "'AV Period Average'" not in text:
        anchor = "\n\tcolumn Period\n"
        assert anchor in text, "column Period not found in Trend.tmdl"
        text = text.replace(anchor, measures_tmdl() + anchor, 1)
        with open(TREND_TMDL, "w", encoding="utf-8", newline="\n") as f:
            f.write(text)

    dst = os.path.join(CL_PAGES, NEW_PAGE)
    if os.path.exists(dst):
        shutil.rmtree(dst)
    for root, _, files in os.walk(SRC_PAGE):
        for fn in files:
            src = os.path.join(root, fn)
            out = os.path.join(dst, os.path.relpath(src, SRC_PAGE))
            os.makedirs(os.path.dirname(out), exist_ok=True)
            with open(src, encoding="utf-8") as f:
                obj = remap(json.load(f))
            if fn == "page.json":
                obj["name"] = NEW_PAGE
                obj["displayName"] = "Average Variable"
            if obj.get("visual", {}).get("visualType") == "slicer":
                # labels differ between the two models, so start on CL1.PV (CL_Dashboard's TagLabel)
                for g in obj["visual"]["objects"].get("general", []):
                    cond = g["properties"]["filter"]["filter"]["Where"][0]["Condition"]["In"]
                    cond["Values"] = [[{"Literal": {"Value": "'CL1.PV - Cooling Load (kW)'"}}]]
            with open(out, "w", encoding="utf-8", newline="\n") as f:
                json.dump(obj, f, indent=2, ensure_ascii=False)

    pj = os.path.join(CL_PAGES, "pages.json")
    with open(pj, encoding="utf-8") as f:
        pages = json.load(f)
    if NEW_PAGE not in pages["pageOrder"]:
        pages["pageOrder"].append(NEW_PAGE)
    pages["activePageName"] = NEW_PAGE
    with open(pj, "w", encoding="utf-8", newline="\n") as f:
        json.dump(pages, f, indent=2, ensure_ascii=False)

    leftover = []
    for root, _, files in os.walk(dst):
        for fn in files:
            t = open(os.path.join(root, fn), encoding="utf-8").read()
            leftover += [w for w in ("Data1min", "DimVariable", "ChartTitle") if w in t]
    print("pages:", pages["pageOrder"], "| leftover refs:", leftover or "none")


if __name__ == "__main__":
    main()

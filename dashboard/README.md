# MT Chiller 1 – Power BI Dashboards

## Projects in this folder
| Project | Pages | Data |
|---|---|---|
| `CL_Dashboard.pbip` | CL Combined Dashboard, CL Correlation Scatter, All Variables | `20261016_benny - Copy.xlsx` (Hist_Value1–5) |
| `EF_Dashboard.pbip` | EF Combined Dashboard, EF Correlation Scatter | `20261016_benny - Copy.xlsx` (Hist_Value1–5) |
| `summary1mindata.pbip` | Summary 1-min Averages, Variable Explorer | `CL_Correlation_Hist1-5.xlsx`, sheet `Data 1min` |

`CL_Dashboard` and `EF_Dashboard` were split from the former `CombinedTrend.pbip` on 13 Nov 2026. Each has its own copy of the semantic model, so a change in one does not affect the other.

## Open it
1. In Power BI Desktop: **File → Open → Browse** → the `.pbip` file above.
2. If the visuals are empty, click **Home → Refresh**.
3. To share as a single file: **File → Save as → .pbix**.

If Desktop refuses to open the `.pbip`, turn on **File → Options → Preview features →
"Power BI Project (.pbip) save option"** and **"Store semantic model using TMDL format"**, then restart.

## Data (CL_Dashboard and EF_Dashboard)
| Query | Source |
|---|---|
| `Trend` | Sheets `2.Hist Value1`..`2.Hist Value5` of `ExcelFilePath`, with only the 9 tags kept and unpivoted to DateTime / Tag / Value |
| `DimTag` | `data/tag_reference.csv`: tag, description, **client reference value**; step `AvgRef` holds the overall averages used on the Combined pages |
| `DimPeriod` | Built from `Trend`: Hist_Value1..5 with the date range of each one |

- To change the source file, edit the **ExcelFilePath** parameter (Transform data → Manage parameters).
- To change the original reference values, edit `data/tag_reference.csv` and refresh.

## Regenerate
- `python build_summary1mindata.py` rebuilds `summary1mindata` (overwrites layout edits made in Desktop).
- `build_pbip.py` is the original generator of the old `ChillerTrend` project. The dashboards have been edited a lot since, so do not run it – it would not produce the current pages.

# Cable Tracker Setup Cookbook

This document outlines how to turn your BOM files into a self-refreshing Excel tracker using stock Office 365 / Excel 2021 features.

## Folder Structure

| Folder | What goes inside | Why |
|-------|-----------------|-----|
| `01_Source-BOMs` | Each project BOM file exactly as you receive it (`*.xlsx`, `*.csv`, etc.). | Keeps originals intact; Power Query will read from here. |
| `02_Lookups` | Static tables you maintain by hand: `tbl_Vendors`, `tbl_PhaseCodes`. | Reference data rarely changes; easier to curate. |
| `03_Reports` | The final tracker workbook (`Cable-Tracker.xlsm`) and any Power BI files. | Clean separation = good governance. |

Put these folders in a shared OneDrive or SharePoint site so everyone uses the same path.

## 1. Build the Master Data table with Power Query

1. **Get Data** → *From File* → *From Folder* and point to `01_Source-BOMs`.
2. Choose **Combine** in the Combine & Transform window.
3. In Power Query Editor perform the following steps:
   - **Promote Headers** to match your BOM field names.
   - **Detect Data Type** and set key fields like `LINK TO SPEC SHEET` to `Text`.
   - **Trim & clean**: remove trailing spaces using `Table.TransformColumns(..., Text.Trim)`.
   - **Add calculated columns**:
     - `TotalNeed = [ENGINEERING QUANTITY] + [SPARE QUANTITY]`
     - `Variance = [PO QUANTITY ORDERED] - [TotalNeed]`
     - `WeeksToNeed = Duration.Days([DATE MATERIAL NEEDED ON SITE] - DateTime.LocalNow()) / 7`
   - Rename the query to `FactCable`.
4. **Close & Load To** → Data Model (not a sheet).

New BOM files dropped in the folder will load on **Data → Refresh All**.

## 2. Create dimension (lookup) tables

Create these tables in Excel and format them with `Ctrl + T`:

- `tbl_Project` — distinct list of Project Name, in-service date, region.
- `tbl_Material` — MAJOR MATERIAL IDENTIFIER, generic description, cross-section, voltage.
- `tbl_Vendors` — imported from `02_Lookups`.
- `tbl_Date` — calendar table (via Power Pivot).

Create relationships in **Power Pivot → Manage → Diagram View** to form a star schema.

## 3. Key measures in DAX

Example measures:

| Measure | DAX Formula | Purpose |
|---------|-------------|---------|
| `Total LF` | `SUM ( FactCable[TotalNeed] )` | Rolls engineering + spares. |
| `Ordered LF` | `SUM ( FactCable[PO QUANTITY ORDERED] )` | What’s on PO. |
| `Variance LF` | `[Ordered LF] - [Total LF]` | Flags short/over orders. |
| `Weeks to Need (avg)` | `AVERAGE ( FactCable[WeeksToNeed] )` | Prioritise buys. |
| `At-Risk $` | `SUMX ( FactCable, [Variance]*RELATED(tbl_Vendors[UnitCost]) )` | Dollar impact. |

## 4. Build the front-end sheets

- **Control Panel** — slicers for Project, Vendor, Phase Code, and number cards using the measures above.
- **Cable Status** — PivotTable with MAJOR MATERIAL IDENTIFIER rows, Project Name columns, and the key measures.
- **Delivery Gantt** — PivotChart with stacked bars showing required weeks vs PO delivery weeks.
- **Risk Heatmap** — Pivot with Vendor rows, Lead-Time buckets columns, values = Total LF.

All pivots bind to the Data Model, so they refresh via **Data → Refresh All**.

## 5. Optional automation

| Need | Feature | How |
|------|---------|-----|
| Automated folder watch | Power Automate for Desktop | Trigger on new file → move to `01_Source-BOMs` then open tracker & Refresh. |
| One-click PO Shortlist export | VBA or Office Script | Filter `Variance<0` then export as PDF. |
| Data-entry guardrails | Data Validation + dynamic drop-downs | Prevents rogue spellings. |
| What-if slider | Form Controls → Scroll Bar | Tied to a DAX parameter table. |

## 6. Why this approach

- Raw BOM files are preserved and transformations are centralized in Power Query and DAX.
- Scales to large datasets thanks to the column-store engine.
- New columns from future BOM revisions propagate once added to the Power Query schema.
- The model can feed Power BI or Python for advanced analytics.
- Remains familiar to colleagues since it’s still “just Excel.”


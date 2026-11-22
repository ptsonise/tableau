# CLAUDE.md

Context for Claude sessions working on this repo. Read it before answering questions about the script.
Trimmed Nov 2026: this file holds the rules that apply TODAY. Change history, proofs and old numbers are in git
(`git log -p CLAUDE.md`; the full pre-trim version is the commit before "Trim CLAUDE.md").

## What this is
`update_inbound.py` is one script for the raw-material planner (USA plant, solar modules). Inputs (same folder as the script):

| Input | Used for |
|---|---|
| `WSA - Raw Material planning per day.xlsx` | Planner's sheets per material: deliveries (PO, BL, ETD/ETA, status), WSA demand/balance |
| `MRP.xlsx` | Planner's workbook: **Settings** sheet (targets, OEM customers, substitutes, ALLOWED VENDORS, transit lead times). Settings only is valid (`_is_user_mrp`) |
| `MIS RM Dashboard - WSA.xlsx` | MIS: Total Plan Req, Standard BOM, In Transit, RM Status (OEM JW), master data |
| PPC sheet `MB52` (else `MB52.xlsx`) | SAP stock (Unrestricted + Quality Inspection). User 20-Nov-2026: "Find MB52 sheet from PPC from now on" (`ppc_mb52_sheet`, `USER_MB52_SHEET`); MB52.xlsx only when the PPC file has no MB52 sheet. MIS MB52 gives names/keys only |
| `PPC*.xlsx` (optional) | **Module plan** = `WSA Plan` (USA lines 1-4, one column per line; user 19-Nov-2026 "WSA plan is correct"; Wp / model from `Plan` when it exists, else from WSA Plan's own 'Wp' / 'Module Model' columns - PPC of 20-Nov-2026 has NO `Plan` sheet, `_has_ppc_plan`) + `OEM Plan` + `Revised Plan` on top; **qty per module** = `BOM`; physical count = `Count`; `GR` (reference only); `In_Transit`; `PM_CS_Usage` |

Outputs: `MRP_Daily_<date>.xlsx`, `Inbound_<date>.xlsx`, `Daily_Material_Status_<date>.xlsx`, `WSA_Daily_<date>.xlsx`,
`Daily_Email_<date>.eml` / `.html`, `Run_Summary_<date>.txt`, `Shortage_Board_<date>.html`. Run: `python update_inbound.py` (Python 3.9+, pandas, openpyxl, lxml),
~50-70 s. Every run also builds the WSA_Daily file (step 1 = calculation + `align_wsa`, step 2 = normal run on that file, `--no-align`
skips, `--keep-old` keeps older outputs). Last console line: 'Timing: ...'.

**THE PLANNER'S WSA FILE IS NEVER CHANGED** (user 11-Nov-2026, option A). WSA_Daily is a new file each day next to it; older
WSA_Daily files are deleted. `_wsa_tidy` removes script-made leftovers (exact names only; any other '(aligned ...)' file is kept with
a NOTE). If WSA_Daily is locked in Excel: `_put_file` retries, then saves '... (2).xlsx' / '(aligned <date>)' with a boxed WARNING.

**TWO DAILY OUTPUTS - ONE LOGIC (standing rule).** WSA daily (aligned WSA file + Inbound) and MRP_Daily must tell the same story.
Every calculation / mapping rule asked for one is applied to BOTH in the SAME change, without being asked; say in the reply how, and
show both in the before/after comparison. Only presentation choices the user placed on one sheet stay there. If a rule cannot apply to
the other output, say why and ask. Shared logic (one source, check both when changing):

| Logic | WSA daily | MRP_Daily |
|---|---|---|
| Stock = MB52 + PPC count rolled forward, JW at OEM out | Balance on today's row (`align_capture`) | `b['stock']` |
| Demand = module plan x BOM, OEM plan out | `PPC_Plan_Req` SUMIFS | `b['plan']` |
| In transit = PPC In_Transit ONLY, pending qty, plant arrival date (late lines on today) | PPC In_Transit -> column T | PPC In_Transit via `ppc_it_rows` (`ships`) |
| Material of a shipment = your WSA PO + BL# first, then description + supplier + RM Code | `ppc_arrivals` / `read_ppc_in_transit` | `normalize_in_transit` (both use `map_in_transit`) |
| Front / back glass = two materials | one sheet per glass (`_glass_split`) | per material |
| Supplier <-> customer = Settings ALLOWED VENDORS; stock by supplier = MB52 batches / count by supplier | 333 on 2 sheets (`_wsa_shared`) | supplier rows, `allocate_lines` |
| ALLOWED VENDORS on the projection (`vendor_pools`) | AK / AL columns, V subtracts the lock | `_mrp_vendor_pools` -> `x['S']` minus lock |
| ALERT list (`alert_eval`, `alert_why`, `build_alert_rows`, built ONCE) | NO ALERT sheets in Inbound (user 21-Nov-2026: "Daily inbound, delete all sheet alert"; `add_alert_inbound` adds only 'ETD alert') | first tab 'ALERT' (`write_alert_sheet`) |

## Key rules (how the numbers should behave)

### Demand, stock, plan
- **Demand = module plan x qty per module**, by customer and line (not the WSA sheet demand). Plan = PPC `Plan` with `Revised Plan`
  on top (a Revised row replaces the Plan rows of its date + line; names fuzzy-matched; `ppc_plan_sheet()` / `read_revised_plan()`).
  Without a PPC file: `Production Plan` sheet of MRP.xlsx. Qty per module = PPC `BOM` for the customers it lists, else MIS Standard BOM.
  Plan customer -> MRP customer only through `LINKS` / Settings MIS LINK (`link_override()`), never exact spelling.
- **Stock = MB52**, replaced by the PPC `Count` sheet ONLY (newest date column with a number; a 0 there is a real 0), rolled forward
  (count - usage + GR since count). No PPC file = no counts (WARNING). MRP.xlsx / old MRP_Daily 'Stock Count' are never read for counts.
  PPC Count = its MAIN table only (`_count_main_table` cuts a second list). Count sheet layout since 19-Nov-2026: the count column is
  'Physical Count Qty_BOM' (BOM unit; text 'no counting' = not counted) + ONE 'Count date' column (user 19-Nov-2026: "must read Physical
  Count Qty_BOM and Count date"). `_count_main_table` puts Qty_BOM / Quantity into the 'Count Closing Stock (BUN)' column renamed
  'Physical Count Qty' (MRP unit stays: RTV 006 plans in DR, 39,460 KG / 270 = 146 DR; CS / PM use Qty_BOM). 'SAP Closing Stock (BUN)' is not a count.
  Same Key + Qty_BOM + Count date as an earlier row = duplicate, counted once + NOTE (ribbon 036 rows 16/17). A material counted under
  2+ codes rolls forward from the OLDEST count date of the codes with stock > 0 (busbar 047 17-Nov; alias counted 0 on 19-Nov does not reset it).
- **PPC sheet `GR` = REFERENCE ONLY** (user 20-Nov-2026: "GR sheet is for reference only, don't count it in any calculation yet"):
  `read_ppc_gr` -> `GR_REF` (Qty in unit of entry on Posting Date; material via MIS MB52 Final Material Code + alias). Received since count
  stays 'PPC GRN' on 'GR. Date' of PPC In_Transit (`GR_LINES`; was MIS In Transit until 21-Nov-2026); in transit is NOT reduced by SAP GR (tried and reverted 20-Nov: SAP GR is a day
  ahead of MIS, so receipts were counted twice). Daily_Material_Status columns after 'Pending GRN Qty' (BOM Weekly in `DMS_WEEKLY_KEEP`,
  BOM Daily, RM dashboards): `GRR_HEAD` 'SAP GR since count (PPC GR, reference)' (`gr_ref_since`; not counted = '-') and `GRL_HEAD`
  'Last SAP GR (posting date)' (`gr_ref_last`). 333 on 2 sheets shows the material total on both.
- **A receipt always goes to ITS OWN supplier** (user 21-Nov-2026 answer '2'): `read_gr_lines` keeps each PPC GRN line's 'Supplier Name'
  (`GR_SUP`); `estimate_counted_stock` stores the receipts since the count by supplier (`STOCK_EST[m]['rec_sup']`); `apply_count_by_supplier`
  splits only count - usage by the count share and adds every receipt to the supplier row of its supplier (first word match; no row = new
  row). 21-Nov: FLAT invoice 81117080 (5,400 x 312 + 5,400 x 333) -> FLAT only (312 KIBING 51,219; 333 SB Back 51,981 / Subanci Back 42,700).
- **BASE DATE = MB52 file date** (= the PPC file save date when MB52 comes from PPC) when older than today (`base_d0()`): everything is calculated from it, a count dated on / after it is
  taken as on it; an old MB52 = today's stock (no roll-forward). File names keep the run date. Console 'BASE DATE ...'.
- **OEM** (Settings OEM = Yes): OEM plan and **JW stock at the OEM are in NO calculation** ("ignore stock at OEM only"; do NOT remove /
  zero OEM data or change OEM target / plan logic). JW stays visible as data (Action List 'to OEM Stock', Stock Data, No Demand JW,
  Movement Log). Code: the OEM step in `build_new_mrp` subtracts `st_jw` from `b['stock']` once and sets `b['jw_out']=True`; code
  after it must not subtract again; code BEFORE it uses `own_stock(b)`. Supplier 'Make' stock must not read JW (RM Status col L).
  Action List 'Inv. Days at line daily target_with OEM' = (own stock + in transit) / target with OEM.
- **Substitute groups** (Settings SUBSTITUTES): pooled on Action List / MRP Daily / Weekly / MRP by Line / RM Issues; on the ALERT (and
  all built from `ALERT_ROWS`) **one row PER MEMBER** (`solo`, `r['mem']`, `r['grp']`, `r['subnote']`; user: "I want epe 480 epe 420
  separately"). `alert_skip(x)` skips only followers / groups without demand; a group main with no own plan uses the GROUP plan
  (`no_demand`). NOT done (ask first): Settings SUBSTITUTES on the WSA sheets.
- **CS / PM '_Alt' lines = the SAME material** (user 20-Nov-2026: "532 is alternate of 004, separately creates confusion"): PPC BOM 'Line #'
  = '<standard>_Alt' (PMPKM00000532 -> PMPKM00000004) -> `cspm_alt_fold` / `CSPM_ALT`: the alt line is used only when that customer / model
  has no standard line (never on top - it doubled the wrapping-roll demand), its count stock (`cspm_alt_stock`), MB52 and GR since count
  (`estimate_counted_stock`) go to the standard code; one row on every sheet, not offered on Substitute Advice. RM busbars 047 / 048 are
  two materials (both in every module); their Settings SUBSTITUTES (016 / 032, 005 / 033) are already in their stock (PACK_FACTORS alias).
- **Settings ALLOWED VENDORS in every projection** (`vendor_pools`): a customer tied to a vendor uses only that vendor + unknown-vendor
  stock; customers without a rule use everything. A customer with 2+ vendor rows uses its FIRST vendor only. 333 in MRP_Daily takes
  the split from the aligned sheets (`_wsa_sheet_series` / `_vp_from_sheets`). FLAT stock stays with Subanci (user decision).
  **DEAD supply** (user 18-Nov-2026 "why doesn't estimated stock follow?"): stock / arrivals of a vendor NO customer with demand may use
  (312 FLAT) are out from day 1 (`vendor_pools` `dead_st` / `dead_T`, in `lock` = WSA AK; `vp_dead(m)`): ALERT / Material Status /
  dashboards 'Estimated stock today', SAP / count, In transit, Overdue, Next arrival, Arrival at plant, Inv. Days, Action List days, Stock
  Check 'Used in MRP' of that supplier row = 0. WSA Balance C / Plant Days R / Plant + Port Days S stay pooled (planner's formulas; not asked).
- **Tapes: YOUR BOM CODE = THE MATERIAL** (`prepare_mis` `CODE_SWAP`): MIS tape code not in your PPC BOM while the BOM lists ONE other
  roll length of the same tape -> planned under the BOM code (RMTTT00000010 -> 018 x0.5, RMTAP00000025 -> 018 x2.5); MIS plan, In
  Transit and checks converted; the planner's typed WSA rows are already in the BOM code's unit (not converted). Item name = the MIS Common Name
  with YOUR BOM roll length (`swap_name`, set in `build_new_mrp` AFTER the WSA capture - earlier it breaks the alias / sheet link; user 20-Nov-2026:
  MIS names 010 and 018 both 'Cell Fixation Tap-T50' -> shown 'Cell Fixation Tap-100m').
- **Units stay as SAP / MIS / PPC BOM give them** ("keep m2"). WSA label unit / product-type text of another group is corrected to the
  MRP unit / description (`ufix`, `mfix`).
- **Check 'Usage per module: MIS BOM vs your BOM'** (`bom_rate_checks`) compares per customer and uses an alternative code only when the
  main code is not listed. **Check 'Stock vs WSA balance'** compares own stock (no JW).

### Usage/Day and day thresholds
- **Usage/Day** (`x['pl_need']`) = average plan per production day (plan > 0) over the material's TRANSIT LEAD TIME (`usage_day(P, lt_of(b))`);
  OEM plan out, substitute group pooled. 'Inv. Days' = stock / Usage/Day; 'Inv. Days incl. in transit' = (own stock + in transit) / Usage/Day
  (`x['pl_days']`). Names are used everywhere. The 16-week HORIZON of 'below 21 days within 16 weeks' stays.
- **Every day threshold = the material's TRANSIT LEAD TIME** (`tgt_days(lt)` = `lt_used`): target days `b['td']`, ALERT 'under N days'
  / 'Below lead time from' (class labels UNDER LEAD TIME / UNDER LEAD TIME LATER), reorder = Usage/Day x (lt + lt), look-ahead, daily
  report, heatmap colours, RM Issues gap, line threshold `LINE_LT` (longest lead time of its materials). `lt_used` = Settings
  'TRANSIT LEAD TIME' table (`LT_OVERRIDE`, column F 'Your lead time', read from MRP.xlsx Settings) > median ETD -> plant days from the
  WSA deliveries > fallback Settings LEAD_TIME_RM 49 / LEAD_TIME_CSPM 14 (`LT_TABLE` writes the table). WSA Plant Days unchanged.
  Dashboard colours `DAY_BINS(lt)`; Shortage Watch has no Status column (user 14-Nov-2026). Dead code is deleted (user:
  "delete all script that you ignore") - do not keep unused functions / constants.
- **Daily targets and OEM columns = Action List ONLY.** Every other sheet follows the module plan. Action List shows BOTH scenarios
  (do not ask again): daily TARGET ('Below 21 Days From' / 'Days Late' use `x['cw']`, with OEM target) and MODULE PLAN gold band
  (`x['cw_p']` = ALERT 'Below 21 days from' / 'Days late'; copied to MRP by Line). Status follows the plan. Stock-zero periods on the Action
  List = 'Short from -> to (with in transit)' only ('Short From (Stock = 0)' stays on RM Issues).
- **MRP Daily / Weekly**: 'Firm receipt (Open PO, no vessel)' and 'Days of cover (... + Open PO)' rows exist only for materials with an
  Open PO (rowmap `ro` / `rj` = None otherwise - guard every use); Supplier column = the single real supplier (`b['_sup1']`); supplier blocks
  only with 2+ suppliers (`multi_sup`). 'Inv. Days' rows = projected stock at day end / Usage/Day from that day.

### ALERT, Stock Management, Material Status
- **ALERT = first tab of MRP_Daily** (NOT in Inbound since 21-Nov-2026; ALERT (formulas) / Calc / Plan lines / Arrival lines also MRP_Daily only) (`build_alert_rows` -> `alert_eval`, stored in `ALERT_ROWS` / `ALERT_D0`).
  Stock-zero day = S < 0 (stock + MIS in transit - plan; on a WSA sheet = Plant + Port Days (PPC In Transit) below 1). Listed when days
  of stock < REPORT_DAYS (21), or zero days inside the lead time, or below 21 days within 16 weeks (RM AND packaging). Classes:
  STOCK-OUT IN LEAD TIME (red) / STOCK-OUT LATER (orange, RELEASE_WINDOW) / UNDER 21 DAYS (yellow) / UNDER 21 DAYS LATER (blue).
  Columns include 'Below 21 days from', 'Why (cause)' (`alert_why`), 'Days late', 'Arrival at plant (Final AtA)'. Layout (`write_alert_sheet`):
  group band row 1 / header row 2 / data from row 3, 'Sr.' in A, code in C, D 'Item name' (MIS Common Name), E 'Category', 23 columns:
  A-F, STORE G-J (Estimated stock today, SAP stock (MB52), Physical count, Days of stock), PRODUCTION K-L, TRANSIT LEAD TIME M-P,
  BELOW 21 DAYS Q-S, WAREHOUSE T-W; summary boxes, 4 class cards, definitions line below. **No INVENTORY / QUALITY** (reorder / status /
  on hold) on the ALERT or Material Status (the fields are still computed in `alert_stock_fields`). **No 'Action' column on any sheet**
  (ALERT, RM Issues, Check) - do not add action-text columns again. Packaging stays on the ALERT.
- **Stock columns = ONE rule** (`stock_src` / `sup_split` / `alert_stock_src`): 'Estimated stock today' = PPC count rolled forward, not
  counted = SAP (MB52) = the stock every shortage date uses (JW out). 333 per WSA sheet = the suppliers whose estimate adds up to the
  sheet (`_subset_match`). A material with stock from 2+ suppliers = one Stock Check row PER SUPPLIER (`b['sup_sap']`, `b['sup_cnt']`).
- **`item_category(m)`** = ONE rule: PPC Count 'Item Category' / 'Item' (header accepted either way) first, then Material Group col E >
  MIS Standard BOM Understanding Name > MRP.xlsx Material Group col A > '-'.
- **CONSUMABLES = the user's list** (`CONSUMABLE_CODES`, `is_consumable(code)`: RMTAP00000018, RMTAP00000025, RMTTT00000018, RMTTT00000009,
  RMLAB00000097 / 098 / 121, or their BOM alias). Wins over the PPC Count Item: CS group on dashboards, label 'CS - Consumable', OFF RM
  Issues. Other groups by PPC Count Item (`count_items`): Packaging (or RMLAB prefix, `is_pm_like`) = PM, else RM; code not on Count
  takes its alias code's Item. `is_rm` / lead times unchanged.
- **Stock Check**: 'Shortage risk' = the ALERT label (`alert_label(code)`; RM Issues still uses `risk_info`); rows in ALERT order;
  columns Item name, Category, Estimated stock, Lines / Project Name / Usage per module by customer / Module (P..S). 'Received since
  count' = `gr_between(code, count date, today)` from PPC GRN on MIS 'GR. Date' (Excel number -> `_any_date`). The CURRENT MONTH band was
  removed (user) - do not add it back.
- **'ALERT (formulas)' + 'ALERT Calc' + 'ALERT Plan lines' / 'ALERT Arrival lines'** (MRP_Daily only, `write_alert_formulas`, `alert_lines()`):
  the user can recheck the ALERT with Excel formulas; every row is verified against `r['P']` / `r['Tp']` (a row that does not add up gets
  per-day projection lines + console NOTE); 'Check vs ALERT' compares each column with the ALERT's own letters; no cached values
  (fullCalcOnLoad); use the `_xlfn.MINIFS` prefix; the line sheets get no per-cell font (speed).
- **Stock Management ('Material Status' sheet)**: `add_stock_mgmt_dashboard`, rows = EVERY material in AL_ORDER (RM + PM + CS, also No
  demand), right after the ALERT; column F 'UOM' (`_add_uom_col`). The Action List stays in MRP_Daily (built in memory, other sheets use it).
  All MRP_Daily sheets are visible. Self-check skips Check / Stock Data / No Demand / Substitute Advice (not in shortage order).
- **Action List = RM Issues look** (title 20 pt navy, header row 4 light grey, Calibri 10; 'Priority' column = RM Issues rule: High = zero
  inside lead time, Medium = short later, Low = under 21 d, OK, No demand; no stock-source columns - they are on Stock Data).

### Dashboards
- **One risk order on every dashboard** (`bom_risk_key` on `bom_day_cover`; RM dashboards `_risk`): first day the daily cover (stock that
  morning / that day's demand) goes below 1, then below 5, then lowest cover today; no demand last; inside the RM / PM / CS groups.
- **NO dashboard sheets in MRP_Daily** (user 21-Nov-2026: "Daily mrp, delete all dashboard sheet"): both BOM dashboards are still built in MRP_Daily
  (Daily_Material_Status copies them), then deleted from MRP_Daily right after `make_daily_material_status`.
- **MRP by Line** (`make_line_view_xlsx`): NO merged cells - the material columns (Component, Stock, module-plan Usage/Day ...) are
  repeated on every line row (user 21-Nov-2026 "Why Mrp by line empty": a later openpyxl save kept only the first row of a merge, so a line
  filter showed blanks). Dates in the Action List columns get 'dd-mmm-yy'.
  Column D 'Line stops because of' (user 21-Nov-2026) = 'code name' of the material(s) that stop the line on 'Line stops on'
  (`line_run_days` `stop_by`, scenario 't'); Component is column E.
- **BOM Material Dashboard - Weekly** (`HEAT_NAME`, `add_cover_heatmap`, same rows as the ALERT) and **- Daily** (`DAILY_NAME`,
  `_bom_daily_sheet`, built LAST before `unify_workbook_style`): weekly value = the daily cover on that week's date (`bom_day_cover(r)[7w]`,
  colours `DAY_BINS`, short = below 1). Column A Code, B Item name, C Category; E..G Stock / SAP stock / Physical count and next columns
  from Material Status (`HEAT_MS_COLS`, `_heat_ms_fill`; EPE members = own values `own_ms`); 'Short from' in bold red. Final layout
  `_heat_final`: no title / note rows (header row 2, data row 3), 9 pt, columns E..'Days incl. in transit' outline level 1. Groups
  RM -> PM -> CS (`mgrp`, RMLAB checked first); a band row per non-empty group. Colours (MRP_Daily): softer dark (DASH_BG / DASH_TAB
  33383F, DAY_NA 474D57, bands 2B5A84, DASH_RED FF8A80). The all-light theme was rejected for MRP_Daily.
- **Daily_Material_Status = TWO visible sheets** (user 18-Nov-2026, RM Issues added 19-Nov): 'BOM Material Dashboard - Weekly' + 'RM Issues' (MOVED here from MRP_Daily on 19-Nov-2026: copied in `make_daily_material_status`, deleted from MRP_Daily after it; your Owner / Status / Date Closed come from the newest Daily_Material_Status, `load_rmi_prev`) with columns `DMS_WEEKLY_KEEP` only
  (Item name, 'Component Description from BOM' (PPC BOM Description, `BOM_DESC`; user 19-Nov), Item Category, UOM, Estimated stock today, Inv. Days, Usage/Day, First stock-zero day, Stock zero from → to, 'Arrival at plant (Final AtA)' (replaced Next arrival, user 19-Nov)) +
  the week columns (`_bom_weekly_keep`); every other sheet hidden. Header row 2 = centered, wrapped, 42 high, autofilter
  (`_dms_weekly_header`, user 18-Nov-2026 screenshot) - the ONLY wrapped row; data rows stay one line. Shortage Watch / Materials Status / ALERT are built then deleted (15-Nov).
- **'Pending GRN Qty'** (user 18-Nov-2026, DMS ONLY, information - in no calculation): PPC sheet 'Pending GR' (`read_pending_gr`: PO.No,
  Invoice No, Supplier Name) matched to PPC In_Transit lines on PO + Invoice + supplier first word (`pending_grn`); qty = the line's Final
  Pending Qty in MRP unit, material = the traced code (`map_in_transit`; one invoice can be EPE + EVA). Column after 'Arrival at plant (Final AtA)' (BOM
  Weekly, in `DMS_WEEKLY_KEEP`) / after the in transit supplier column (BOM Daily, RM dashboards). Unmatched rows -> console NOTE.
- **Daily_Material_Status** (`make_daily_material_status`, history): sheets 'Shortage Watch' (FIRST tab, user 14-Nov-2026: the ALERT columns without Sr. / Alert + 'Item Category' after Category, same rows and order as the ALERT,
  read from the workbook's ALERT sheet, one-line text, no frozen panes / wrap; `_shortage_watch`), 'Materials Status' (HIDDEN; Direct Material + OTHER MATERIALS band; Other Materials
  = PPC Count materials not in the BOM, need = 'Material Req Qty per Day'), 'ALERT' (HIDDEN, no class cards), BOM Weekly (opens first),
  BOM Daily (hidden here), 'RM Material Dashboard' (hidden) + 'RM Material Dashboard - MIS' (`DASH_MIS`, the WSA Dashboard read from
  column V 'Plant + Port Days (PPC In Transit)' found BY HEADER `vcol`; no fallback: no header = sheet not built + NOTE 'close the WSA file
  in Excel'). `_wsa_dashboard_to` works the WSA Dashboard VLOOKUPs out in Python (WSA cached values are stale). RM dashboards
  (`_dark_rm_dashboard`): A Code(s), B Item name(s), C Planning sheet, D Category, E Today's Morning stock (before today's arrivals), F
  supplier, G 'Short from', days from H. Every dashboard has Code + Item name; 'In transit supplier name (qty)' column after 'In transit'
  (`_add_it_sup_col`, DMS only). Both Material Status / ALERT copies: `_remap_rows` (no group band, header row 1). Consumable Unit = PPC sheet `PM_CS_Usage`, or `CS_Usage` when the planner renamed it (Nov 2026; column 'Code' + 'Consumable Unit ...'). Usage-per-module column at
  the end of both Status sheets; Direct Material has no STORE / WAREHOUSE / PRODUCTION cards (MRP_Daily keeps them).
  **Look of DMS dashboards = the user's 'WSA Material Coverage' page** (`_coverage_light`, DAILY FILE ONLY): white page, THREE cover colours
  (red under 7 days `COV_RED_DAYS`, amber under 21 days, green covered, blank = no plan / na); then `_tidy_dashboard`: NO freeze panes, NO
  wrapped text, every row 15 high except header row 4, one width per day column, supplier column capped at 48. Do not add frozen panes /
  wrap back. (`_fix_views` in `_safe_save` drops selections of panes that do not exist - Excel 'Repaired Records'.)
- `sort_wsa_dashboard()` re-orders the WSA_Daily file's own Dashboard rows at XML level in the RM dashboard order (`DASH_MAP`); formula
  refs follow the row; on a base-date run F1 = DATE(base) instead of TODAY().

### WSA_Daily file (aligned planning sheets, `align_wsa`)
- Every visible planning block -> one sheet in ONE layout (`WSA_FIX1` + `WSA_SLOTS`=8 customer columns + `WSA_FIX2`): rows 1-6 info (A2
  stock, A3 'Supplier & Material', A4 unit, A5 codes, A6 totals), header row `WSA_HDR`=7, one row per day from the 1st of the current month
  (`keep_from`) to 31-Mar-27. Columns A..E Date / Day / Balance / Planned / Demand; F..M customers; N..S Actual modules, Actual, MW, Actual
  MW, Plant Days, Plant + Port Days (S); T 'PPC In Transit (arrivals)', U supplier name, V 'Plant + Port Days (PPC In Transit)';
  W Delivered, X Scheduled Delivery .. AF Vessel ETA, AG Status, AH Invoice, AI Reference, AJ Remarks 2, AK 'Vendor rule: not usable', AL own
  balance per vendor (by header: `XL(k)`; with 9+ customers V moves right). Opens at today's row, frozen D8, autofilter row 7, red < 0.
  Two-block sheets split (`WSA_SPLIT_NAMES`); 'Stock' and 'Total_Plan_Req' sheets dropped (`WSA_DROP_SHEETS`); recognised by
  Total_Plan_Req OR PPC_Plan_Req (`_WSA_MARK`); no Total_Plan_Req -> plan end / names / codes / UOM from MIS 'Total Plan Req'.
- Rows before today = copied values; today and later = formulas. Balance today = typed own stock; C(r) = C(r-1) - Actual(r-1) + Delivered(r-1).
  D / E / customers / MW = SUMIFS on `PPC_Plan_Req`. Plant Days = C/E. Plant + Port Days = (C + Scheduled row 2 + SUMPRODUCT(open
  Scheduled, remarks without 'canc', Delivered blank)) / E. V = IFERROR((C + SUM(T today..r)) / E, "na"). X2 = SUMIFS(Archive, State 'Not
  delivered'). Archive sheet accumulates (deduped). Every formula has its cached value (mirrored in Python); calcChain removed; before
  replacing anything the file must read back (ZIP + XML parse) or the run reports 'ERROR aligning ... nothing was replaced'.
- **Front / back glass on their own sheets** (`_glass_split`, side by `_glass_side`): '<sheet> Front' + '<sheet> Back' (280 / 268; SB
  312 / 333; Subanci 356 / 333); quantities split by the BL#'s PPC In_Transit lines (`_it_bl_index`), else 1:1; marker in A1 = never split
  again. **One material on 2+ sheets** (`_wsa_shared`, today back glass 333): each customer on its own sheet (customers = labels with
  numbers from the 1st of the month; supplier goes to the sheet whose customers have it as FIRST allowed vendor; other stock -> any-vendor
  customers (Encap), else first sheet; PPC arrivals by the PO + BL# the planner typed if allowed, else by supplier). An aligned file of
  the older merged version is detected by its A1 notes -> WARNING: run on the planner's own WSA file.
- **PPC In Transit** (`read_ppc_in_transit`, `ppc_arrivals`): reads PPC sheet `In_Transit` (stops after 500 empty rows); qty = Final Pending
  Qty (else Final Qty - GRN), date = Final ATA2 / Final AtA. Robust reading (`_it_sheet` name variants, header row = first row with 2+
  known headers, text numbers / dates). When nothing arrives `ppc['why']` explains in T3 of every sheet, Inbound 'PPC In Transit' and a
  last WARNING line. `_hdr_role` 'ppcsup' is checked before 'ppcit'.
- **MATERIAL of a shipment = traced to your BOM** (`map_in_transit`, `in_transit_catalog`): your WSA FIRST (`wsa_ref_index` / `wsa_find`:
  PO + HBLNo./MBL suffix match > BL > PO; codes of the matched sheets that fit the description decide), else scoring by description + supplier
  (`SUPPLIER_ALIAS`) + RM Code evidence; Final RM Code replaced only when another code is 2+ points better; 'not a BOM material' for weak
  matches; CHECK (red) only when the WSA sheet has the same part in another size / HS. MRP_Daily uses the same tracing (`IT_FIX` -> Check
  'MIS Final RM Code does not match the shipment - traced code used'). Inbound sheet 'PPC In Transit' = the trace of every line.
  **Copied description = the unit price decides** (user 21-Nov-2026 'go'): 2+ lines of ONE invoice with the SAME description but
  DIFFERENT 'Unit Rate', each price with its own Final RM Code -> Final RM Code used, before the WSA / description step (`map_in_transit`
  step 0, `IT_COPIED`, console NOTE). 21-Nov: Kibing KB-WSA-260716B-5 / -6 (both lines 'HTAR', 4.02 = 312 front, 4.07 = 333 back) ->
  36,000 front + 36,000 back (was 72,000 front).
- **IN TRANSIT = PPC In_Transit ONLY** (user 21-Nov-2026: "Script find in transit from PPC only. Ignore mis in transit"): `_CachedWB`
  hands every reader of the MIS 'In Transit' sheet the PPC 'In_Transit' rows (`ppc_it_rows`, same columns as MIS): MRP_Daily projection,
  ALERT, dashboards, email, Check, In Transit Status, received since count (PPC GRN / GR. Date), glass split by BL#, ETD alert. Same rules as
  WSA column T (`read_ppc_in_transit`, mirrored exactly; date = `_ppc_it_date` for both): Final ATA2 ONLY when Final AtA holds a date too
  (PPC fills ATA2 with the file date when AtA is blank), else Revised ETA Port + 5; NO date = the pending line is NOT counted until a date
  is entered (user 21-Nov-2026 answer '1c'; console NOTE lists them - 21-Nov: RTV 006 air invoices HT-WRE-20261113-15..18, 4 x 16 DR
  'At ORIGIN'); quantity = Final Pending Qty., else Final Qty. - PPC GRN (text numbers read).
  **A pending line counts only with ETD + ETA (Revised ETA Port, else ETA as Per Booking) + BL# (HBLNo. or MBL) + Final AtA** (user 21-Nov-2026:
  "must be enough information: ETD, ETA, shipping documents, ATA"; `_ppc_it_missing` / `PPC_IT_NEED`, same gate in `ppc_it_rows` and
  `read_ppc_in_transit`; missing lines in the console NOTE with the missing fields; a column missing from the sheet = not checked + WARNING).
  Lines with ETD / ATD after the file date (not departed, 'AT ORIGIN') still count when all 4 are filled (user 21-Nov-2026: "Keep them in" - do not ask again).
  No PPC In_Transit = no in transit + WARNING (MIS is NOT a fallback). Console 'IN TRANSIT = <PPC file> sheet In_Transit ONLY (N lines ...)'.
  Labels say 'PPC In_Transit' (sheet name 'RM Material Dashboard - MIS' kept). Function / variable names with 'mis' stay (MISBL,
  `_mis_date_of`, `build_mis_intransit`): they now read PPC rows. MIS is still read for plan, BOM, master data, MB52 names, RM Status.
- **PPC In_Transit dates are the latest**: the planner's typed WSA rows can be older and the script does NOT move them (offer only if
  asked): MRP_Daily Check 'Same shipment, different arrival date' ("PPC In_Transit is the latest: move the WSA row to dd-mmm"); Inbound
  column AQ + 'To chase' reason when PPC is 3+ days off (`_mis_date_of`, suffix match).
- **Other BOM materials** (no WSA sheet): `CAP[0]` = `align_capture`; `other_materials_calc()` + `add_other_materials()` -> Inbound sheet
  'Other materials' + 'Below zero' rows. Solar cells are not in the MRP. Pine wood pallets stock 0 = a real count.
- **Inbound 'Below zero'** (`below_zero_periods()`): periods where Plant Days < 0 / Plant + Port Days < 0; Summary R/S + sheet. Glass 333
  deliveries of both sheets = 'G12R Back Glass' (`grp_mats`, `name2grps`, `pool_g`; MRP_Daily Check `_sheet_pools`).
- **ETD alert** (`ETD_LATE`, `write_etd_late_sheet`, sheet 'ETD alert' in BOTH Inbound and MRP_Daily): every open WSA delivery row (not
  cancelled; + Archive 'Not delivered') with an ETD whose BL# is not on PPC In_Transit. `_etd_flag`: ARRIVES AFTER STOCK-OUT > ETD PASSED
  > SHIPS WITHIN `ETD_SOON` (7) DAYS > ON PLAN. A date in the BL# column = no BL# (`_bl_txt`). In no projection.

### Daily email (`write_daily_email`, `Daily_Email_<date>.eml` + `.html`) - the email the user sends
Layout = the user's 18-Nov-2026 sample: Dear all, intro (sources + base date), 'Needs attention' (2 bullets, NO reschedule sentence - duplicate, user 18-Nov),
lead-in line, '1. Critical' + one line ('Materials will run out of stock within the transit lead time. ' + reschedule sentence), sub-tables
'1.1 Raw Materials' / '1.2 Packaging' / '1.3 Consumables' (columns Material / Supplier / Cover today / 'Stock zero from → to (⚠ = starts
inside the transit lead time)' = ALERT `r['per']` in red, no wrap / 'Arrival at plant (Final AtA)' = ALERT `r['ata']` / Situation), '2. Upcoming ETD Alert' + one line
+ '2.x' tables Material / Supplier / Deliveries / Quantity / Next ETD / WSA plant date (numbers follow the sections / groups that have rows), Notes. NO signature (user 18-Nov-2026: Outlook adds it). Aptos 9 pt, no emoji. Tables split into Raw material / Packaging / Consumables (empty ones not printed; `is_pk`
= the `mgrp` rule). NO Summary section, NO counts in the subject / text, NO Overdue / Order-this-week tables, NO 'Alert' / 'Runs out'
columns on the ETD table; ETD rows = FUTURE ETDs that ARRIVE AFTER THE STOCK-OUT only, one line per material; facts only (no advice text).
Supplier (`_email_suppliers`) = in-transit suppliers (333 per sheet) > PPC BOM supplier (alias codes too) > MRP supplier-row stock
suppliers (glass / 2+ suppliers) > WSA Shipper > 'no supplier on file'. To / Cc = `Daily_Email_Settings.txt`
(`_email_settings`; company data - never in the script or the repo). Daily_Material_Status is attached; X-Unsent = Outlook opens a draft.

- **Material_Shortage_<date>.html** (`write_html_dashboard`; renamed from Material_Dashboard 21-Nov-2026, older files of both names deleted; user 21-Nov-2026 "dashboard like picture, not in Excel"): sorted by Inv. Days today (no in transit) ascending inside RM / PM / CS; 'Passed, not received' line always visible on the row (not in Details); NO material code on the page (user 21-Nov-2026; search still finds codes); under the name 'Usage/Day N UOM' on the row (not in Details), NO lead-time text (user 21-Nov-2026; runway tick stays); BOM Weekly as a dark web page
  for phone / browser; same rows / groups / order (`_heat_expand`, `_heat_grp`, `bom_risk_key`), cell = `bom_day_cover(r)[7w]`; FEW COLOURS (user 21-Nov-2026 "too colorful"):
  red fill < 7 (incl. < 1; user 21-Nov-2026 "1-7 day is also red", no amber), grey outline 7-21, plain grey number 21+, white underline = lead-time week; red dot / red runway bar only
  for stock-out inside the lead time (else no dot / grey bar); Inv. Days badge red < 7 / grey; 'Next arrival <date> · qty' line ON the row (user 21-Nov-2026 "unhide next arrival"; not in Details); no page gradient; material column 380 px (phone = half the screen, `--mw`); WEEKS WITHIN THE LEAD TIME ONLY (user 21-Nov-2026: columns up to the longest lead-time week of the listed rows, each row blank after its own lead-time week; max `DASH_HTML_WEEKS`=3, user 21-Nov-2026 "Make it show 3 weeks"); ONLY materials with Inv. Days today (`r['days']`) <= 21 (`DASH_HTML_MAX_DAYS`; user 21-Nov-2026; no demand = not listed); badge 'Inv. Days 1.8' (stock only, no in transit) on every row, header 'Material · Inv. Days today', 'Details ▾' on every row + hint line above the table (user: people do not know to tap), 2 separate detail fields; ring '14 of 50 at <= 21 inventory days' kept, NO KPI class cards (user 21-Nov-2026: cards deleted, 'keep the ring'), RM/PM/CS chips, search, runway bar (days to stock-zero vs lead time), tap = details. Presentation only; older ones deleted.
- **Shortage_Board_<date>.html** (`write_shortage_board`, called right after `write_html_dashboard`; user 21-Nov-2026 "Create new web dashboard
  for shortage, I want new design"): a SECOND page next to Material_Shortage (that page is unchanged). Same rows (Inv. Days today <= 21, no
  demand out), ALERT fields only. One card per material, sections 'Runs out inside the lead time' (class 0) / 'Low cover', highest risk first
  (`bom_risk_key`, no RM / PM / CS grouping; RM / PM / CS chips + search filter). Card = name + group tag, big Inv. Days today (red < 7), Usage/Day /
  stock / lead time, DAY STRIP from the base date up to the longest lead time of the listed rows (`BOARD_MAX_DAYS` cap; faded after the row's own lead
  time; white line = lead time): box = `bom_day_cover` (red < 1 = stock zero, dark red < 7, grey 7-21, dark 21+, blank = no demand that day), blue
  triangles = ALERT `Tp` arrivals; stock-zero periods (`per`) + next arrival -> Inv. Days; tap = details.
  BOM USAGE on the card, NO LEAD TIME (user 21-Nov-2026 "show bom usage, hide lead time"): line 'BOM usage N UOM / module' (one value, or
  'Customer N' per customer when they differ) from the Action List `usage_src` (same text as 'Usage per module by customer'; substitute '+'
  lines left out), full list with source in Details; no 'lead time N d' text, no lead-time line / legend (boxes after it stay faded). Dark default, sun button (same
  'web_theme' key), own CSS (not `_WEB_CSS`). Older ones deleted. Console 'SHORTAGE BOARD: ...'. Presentation only.
- **WSA_Daily_<date>.html** (`wsa_html_sheets` -> `wsa_alert_match` -> `write_wsa_html`; user 21-Nov-2026 "Create web dashboard
  for daily wsa output", then "Do all" = line it up with Material_Shortage): written in STEP 2 next to `write_html_dashboard` (needs the ALERT
  rows) when the run reads a WSA_Daily file (no aligned sheets = no page); older ones deleted; console 'WSA WEB PAGE: ...'. One row per
  planning sheet read from WSA_Daily's CACHED values (columns by header). Lined up with Material_Shortage: name = MIS Item name (sheet name +
  code under it), badge = Inv. Days today, stock = Estimated stock today (usable, dead FLAT out; WSA Balance in Details), order = RM -> PM ->
  CS, lowest Inv. Days first, `bom_risk_key` (sheets without an ALERT row last), lead-time week underlined. Sheet -> ALERT row: `where`
  (333 per sheet), else the sheet's codes incl. MRP code (`PACK_FACTORS`: busbar 032 -> 047) and BOM alias both ways (tape 018 <-> 025).
  Week cells = column V on day 7w (`WSA_HTML_WEEKS`=3 weeks, user 21-Nov-2026), chart C vs C + cumulative T - AK (= V x E), PPC arrivals, open WSA deliveries. Note lists
  Material_Shortage materials with no WSA sheet (barcode 121, flux 1341, pallet 535 today).
  NO filter buttons (user 21-Nov-2026: 'Stock-zero ahead' / 'With PPC arrivals' removed; search only) and NO count cards ('stock-zero within
  14 days' / 'stock-zero ahead' / 'PPC arrival days' removed - do not count stock-zero sheets). Instead a full-width panel 'within 21 days'
  under the title (`#w21`, user 21-Nov-2026 "name all materials"): every sheet with Inv. Days today <= `DASH_HTML_MAX_DAYS`, page order,
  NO groups, highest risk first (`rk`, `bom_risk_key`; user 21-Nov-2026, both panels), columns Material (MIS name + sheet) / Inv. Days (stock only, `days`) / Inv. Days incl. in transit (ALERT `days_it`) / Next PPC arrival within
  21 days = the sheet's own column T (`a21`, base date .. base + 20; late lines sit on the base date) with '+N more, total'; tap = opens the sheet row.
  SAME panel on Material_Shortage (user 21-Nov-2026 "do same card for material web page"), under the title, ring kept: every listed material,
  second line = group (+ WSA sheet for 333), arrivals = the ALERT row's `Tp` (projection arrivals, dead FLAT supply out) first 21 days; tap
  resets the RM/PM/CS chip to All. ONE code: `w21Panel` in `_WEB_JS`, CSS `.w21` in `_WEB_CSS`.

- **Inv. Days incl. NEXT ARRIVAL ONLY** (user 21-Nov-2026: "Web page and pictures, I also want to see inv day with the next arrival only"):
  `alert_eval` `days_nx` = (Estimated stock today + qty of the first arrival in T, late lines on the base date) / Usage/Day. Shown on
  Material_Shortage (row 'Next arrival ... → Inv. Days N', Details, 'within 21 days' panel column), WSA_Daily.html (row, Details, panel;
  from the sheet's ALERT row) and the chat pictures / .txt ('next <date · qty> → N d', in transit on its own line, card 102 px).
  Presentation only: in no Excel output and no calculation.
- **Both web pages: material column = 2/3 OF THE SCREEN, week columns = 1/3** (user 21-Nov-2026 "1st column is bigger like 2/3 of the
  screen"; was half / half): `--mw` = calc(66.6vw - page padding), JS `setWC` sets `--wc` = max(34 px, window width / 3 / number of weeks) on load and resize; table not stretched to full width (`.wrap`
  fit-content); Details panel capped at the table width (`--mw` + `--wc` x `--nw`) so opening it does not widen the material column.
  34 px minimum per week.
- **Both web pages: each row = its own card with a small gap** (user 21-Nov-2026 "make each row a small gap for easier viewing"):
  `border-spacing:0 3px` (was 6 px, made compact 21-Nov-2026), rounded row cards, hover / opened row lighter with a blue left edge, group names as plain headings, Details as its
  own card. CSS only.

- **Both web pages: ONE shared look** (user 21-Nov-2026 "more clean and clear and comfortable for viewers"): `_WEB_CSS` / `_WEB_JS` /
  `_WEB_TOP` / `_WEB_SRCH` are used by `write_html_dashboard` AND `write_wsa_html` - change the look there once. Dark by default; a
  sun button switches light / dark (remembered per browser, `localStorage` 'web_theme'); segmented filter chips with counts; search
  icon; legend as small pills (scrolls sideways on a phone); week header day bold, base-date week blue; rows keyboard-friendly (Tab +
  Enter); Details = label / value cards, 'Arrival at plant' one arrival per line, stock-zero dates red; WSA detail tables plain lines.
  Presentation only.
- **Both web pages: COMPACT, PHONE FIRST** (user 21-Nov-2026: "Dark by default. The light/dark button stays. Phone first ... Compact:
  less padding, smaller cards, and more rows on the screen"): `_WEB_CSS` base rules = phone width, `@media(min-width:700px)` = desktop.
  Body 18.5 px (user 21-Nov-2026 "Web page, make bigger fonts size" + "more bigger 20%": every font x1.44, phone search box stays 16 px),
  material column 2/3 of the screen (badge under the name on a phone), week cells 30 px high (max 64 px wide), names up to 2 lines, WSA sheet / code line 1 line (ellipsis, full
  text on hover / search), 'within 21 days' panel scrolls inside itself (34 % of the screen on a phone, 42 % on a desktop, one line per
  material on a desktop), legend one sideways-scrolling line on a phone, search 16 px on a phone (no iOS zoom). Rows per screen
  21-Nov run: Material_Shortage phone 7 -> 10, desktop 8 -> 12; WSA_Daily phone 5 -> 8, desktop 6 -> 11. Week width rule unchanged.

### Housekeeping
- **Old outputs are deleted** (`delete_old_outputs`, only after a full run saved today's files): earlier Daily_Material_Status, Run_Summary,
  ALL earlier Inbound files, MRP_Daily, emails, WSA_Daily. Exact script names only; `--keep-old` keeps all.
- `_WDT_CACHE` (WSA Dashboard read once) only inside the daily-file build and closed after (`_wdt_close`; open handles block the WSA_Daily
  replace on Windows).

## Numbers register (added 21-Nov-2026; enforced by `check_rules.py`)
Every hard-coded threshold / fallback has an OWNER. **USER** = the planner decided it: never change it without asking first.
**DEFAULT** = nobody confirmed it: do not change it silently, do not quote it as the planner's rule, and ask the planner to confirm it
when it matters. A number NOT in this table = treat as DEFAULT and add it here (and to `check_rules.py`) when you touch it.
After agreeing a change: update the code, this table and `check_rules.py` in the SAME commit. `python check_rules.py` must PASS.

| Number | Value | Owner | Note |
|---|---|---|---|
| Every day threshold = the material's transit lead time | rule | USER (Nov 2026) | target days, ALERT 'under lead time', reorder, look-ahead |
| `COV_RED_DAYS` red cover | 7 days | USER 12-Nov | "Red is under 7 days" |
| Web pages: `DASH_HTML_MAX_DAYS` / weeks (`DASH_HTML_WEEKS`, `WSA_HTML_WEEKS`) / colours / width | 21 / 3 / red <7, 7-21, 21+ / weeks 1/3 of the screen, min 34 px | USER 21-Nov | both pages (WSA web page too) |
| `BOARD_MAX_DAYS` Shortage_Board day strip | 63 days | DEFAULT | strip = longest lead time of the listed rows, capped here |
| `CONSUMABLE_CODES` | 7 codes | USER | |
| `LEAD_TIME_RM` fallback | 49 | USER 21-Nov | only when the Settings row is missing (Settings row wins) |
| `LEAD_TIME_CSPM` fallback | 14 | USER 21-Nov | |
| `PLAN_AVG_DAYS` Action List daily plan | 49 days | USER 21-Nov | fixed in the code, NOT a Settings row |
| `CUST_TARGET_DEFAULT` | Pepper 3000, Lucky 7 3000, Crimson 3600, SB 3600, Snowflake 4000 OEM | USER 21-Nov | only when Settings has no CUSTOMER DAILY TARGET |
| `ALERT_HORIZON` / `RISK_DAYS` / `RELEASE_WINDOW` / `ETD_SOON` | 112 / 5 / 7 / 7 | USER 21-Nov | Settings rows |
| `RESCHED_OUT_X` / `SHORT_HORIZON` / `ATA_SHOW` / `MOVE_DAYS` | 3 / 60 / 4 / 14 | USER 21-Nov | |
| `SHORT_LEAD` / `LONG_LEAD` | 20 / 75 days | USER 21-Nov | |
| `DASH_STEP2` / `DASH_BLUE1` / `DASH_BLUE2` / `DASH_MIN_LT` | 14 / 14 / 39 / 14 | USER 21-Nov | MRP_Daily dark colours |
| `IT_STATUS_ONLY` / `UNLIMITED_CODES` | False / PMPKM338, 523 | USER 21-Nov | |
| PPC line without ATA | Revised ETA Port + 5 days | USER 21-Nov | since the ETD+ETA+BL#+AtA rule: only when the sheet has NO Final AtA column |
| PPC In_Transit line counts only with | ETD + ETA + BL# + Final AtA | USER 21-Nov | `PPC_IT_NEED`; STATUS text still not used |
| Status WATCH / line CRITICAL | lead time + 14 / under 2/3 of lead time | USER 21-Nov | |
| Tracing / Inbound AQ / In_Transit read stop | 2+ points better / 3+ days off / 500 empty rows | USER 21-Nov | |

Old Settings rows (target days A/B/C/packaging/Z, line CRITICAL, daily report, look-ahead, safety, days of cover, weeks) are NOT read;
user 21-Nov-2026: "delete rows" -> removed from MRP.xlsx Settings (cleaned copy sent); the script does not write them.

## Known data quirks (NOT script bugs)
- MB52.xlsx has no save date inside -> the file date = the Windows file date (an uploaded copy gets today's date).
- WSA sheet `Glass G12R Topcon Subanci` (planner's original) lists only 356 but its Demand / deliveries count front + back glass (MIS: 356
  + 333); on the planner's original file the Check 'Plan' mismatch for glass is expected (the aligned file splits it).
- 268 = back glass (printed, with hole), 280 = front (HTAR) per MIS Common Name. SAP numbers the user quotes can come from newer files.
- Customer spelling `Clenera Crimson` vs `Cleanera Crimson`: fix it on the Settings sheet.
- "In PPC In_Transit, not in WSA" shipments on the Check sheet = the planner has not added them to the WSA sheets yet.
- Daily_Material_Status 'Other Materials': PPC Count column K ('Material Req Qty per Day') is an XLOOKUP into the PPC 'Usage' table and has
  numbers only for BOM materials, so most rows show 'No demand'.

## How to work on this repo
1. **Get a concrete symptom first**: material code, sheet, value shown, value expected. Don't guess at fixes.
2. **Reproduce with the real input files** (the user uploads them; they are not in the repo).
3. **Prove the fix**: run the script before and after, then compare every sheet of the outputs. Only the expected cells may change.
   The console output should stay the same unless the fix is meant to change it.
4. Keep fixes small and explain them in plain language, with a before/after table for the affected materials.
5. **Never commit the Excel input or output files** (company data).
   **Python 3.9+**: no quotes of the same kind inside an f-string expression (3.12+ only). Check every change with
   `/usr/bin/python3.11 -c "import ast;ast.parse(open('update_inbound.py').read())"`, and run `python check_rules.py` (must PASS;
   a FAIL = a registered number moved - see 'Numbers register').
6. **Independent check (user rule, Nov 2026, lightened 14-Nov-2026 to save usage)**: run the `output-checker` subagent
   (`.claude/agents/output-checker.md`; tell it the before-script, e.g. `HEAD~1`) ONLY for LOGIC changes (a calculation, mapping, projection,
   date, stock, plan or vendor rule). Do NOT run it for text, label, colour, layout or width changes - for those do step 3 only. Fix every
   proven finding and re-run the checker, at most 3 rounds. Send the files only on PASS, or say plainly which findings are still open.
   Show the checker's verdict and findings table in the reply. State in the reply whether the checker was run and why / why not.

## Working with the user
- Correct the grammar and vocabulary in the user's message first (US English), then restate the request as a developer requirement, then answer.
- **Two daily outputs, one logic** (top of this file): prove both outputs in the comparison.
- **After every change to `update_inbound.py`, send the updated file to the user (attachment) in the same reply**, after it is tested, with one
  line on what changed. Also send the output files (`MRP_Daily_<date>.xlsx`, `Inbound_<date>.xlsx`) from a fresh run on the latest inputs,
  put in `output/` (git-ignored - never commit or push them).
- Use simple, clear language. Planning terms (MRP, MB52, JW, OEM, BL, ETA) are fine.
- **Say when a new session is better** (standing rule): at the end of a reply, with the reason, when the conversation was already summarized
  once and the task is done; a big new feature is next on a long context; or the next request is a new topic / a new day's run. Handover:
  the work is on the session branch until merged into `master` (a new session starts from master, so offer a pull request; create it only
  when asked), the 5 input files to upload (the planner's own WSA file - the script never changes it), and the open points below.
- The user wants fewer messages per result: when a request has open choices, build the recommended option and say so, instead of asking
  (ask only when the answer changes the numbers).

## Open points (ask the user; remove when answered)
- GR posted in SAP (PPC GR) while PPC / MIS In_Transit still shows it pending = counted twice (20-Nov-2026: invoice 81117083,
  3,600 front 312 + 3,600 back 333, AtA 20-Nov). Offered: drop pending in-transit lines whose invoice is on the GR sheet. Not done until asked.
- SPEED (user, 12-Nov-2026: "The script is too slow"): 68-74 s on the real files (step 1 11-13 / aligned WSA 17-20 / step 2 39-43); no single
  hotspot (30 saves + 28 loads in step 2, `make_daily_material_status` ~1/3). Real gains need (a) dropping outputs / sheets the user does not
  open, or (b) a calc / presentation split. Asked which outputs they open daily + the 'Timing:' line from their PC (not answered).
- PPC Count second list at row 170 (24 packaging rows, stock units): newer count or scratch list? Ignored until answered; if newer it
  must be typed into the main table (pallets 535 / 447, 620, 603, 506 would change).
- Open (Low): RM Issues 'Project' blank for the EPE 480 + 420 group; Stock Management KEY RISK counts that group twice.

- **WEB PAGES = LIGHT THEME ONLY, WAAREE LOOK** (user 22-Nov-2026: "Keep light theme only. No more dark"; "Delete png, no need"): Material_Shortage / WSA_Daily / Shortage_Board use Poppins, brand green (#2B7A3E, light page #F4F8F2), yellow #F5A623 accent, gradient bar + WAAREE wordmark; NO light/dark button, NO dark CSS. Chat_Update pictures / text were removed (`write_chat_update` deleted).

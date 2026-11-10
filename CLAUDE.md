# CLAUDE.md

Context for Claude sessions working on this repo. Read it before answering questions about the script.

## What this is
`update_inbound.py` is one script for the raw-material planner (USA plant, solar modules). It reads four Excel files and writes two:

| Input (same folder as the script) | Used for |
|---|---|
| `WSA - Raw Material planning per day.xlsx` | Planner's sheets per material: deliveries (PO, BL, ETD/ETA, status), WSA demand/balance |
| `MRP.xlsx` | Planner's own workbook: **Settings** sheet (targets, OEM customers, substitutes), BOM, Stock Count |
| `MIS RM Dashboard - WSA.xlsx` | MIS: production plan (Total Plan Req), Standard BOM, In Transit, RM Status for OEM JW, master data |
| `MB52.xlsx` | SAP stock export. **Stock quantities come from here** (Unrestricted + Quality Inspection); MIS MB52 only gives names/keys |
| `PPC*.xlsx` (optional, e.g. `PPC_with_Usage_Shortage_Alert.xlsx`) | **Module plan** = sheet `Plan` + sheet `Revised Plan`; **quantity per module** = sheet `BOM`; physical count = sheet `Count` |

Outputs: `MRP_Daily_<date>.xlsx` (Action List, MRP Daily, Stock Check, Check, ...), `Inbound_<date>.xlsx` (deliveries, Below zero,
To chase, By Week), `Daily_Material_Status_<date>.xlsx` and the WSA file itself, rebuilt IN PLACE (redesigned planning sheets, see
below; ONE WSA file since Nov 2026 - user: "Why i get 2 wsa output file. I need 1 only", answer 'option 2').

**THE PROJECT HAS TWO DAILY OUTPUTS - ONE LOGIC (standing user rule, Nov 2026: "logic I apply for wsa_daily is also for
mrp_daily, this project requires 2 daily outputs").** WSA daily = the aligned WSA file (+ its Inbound file); MRP daily =
`MRP_Daily_<date>.xlsx`. Every calculation / mapping rule the user asks for on the WSA daily is ALSO applied to MRP_Daily, in the
SAME change (and keep the reverse true: the two must tell the same story). Do not wait to be asked; say in the reply how it was
applied to both, and show both in the before/after comparison. Only presentation choices the user placed on one sheet stay there
(e.g. daily targets / OEM columns = Action List only). If a rule cannot apply to the other output, say why and ask.
Shared today (one source, check both when changing):

| Logic | WSA daily | MRP_Daily |
|---|---|---|
| Stock = MB52 + PPC count rolled forward, JW at OEM out | Balance on today's row (`align_capture` after the OEM step) | `b['stock']` |
| Demand = module plan (PPC Plan + Revised Plan) x BOM, OEM plan out | `PPC_Plan_Req` SUMIFS | `b['plan']` |
| In transit = the same shipments, pending qty, plant arrival date; late lines counted on today | PPC In_Transit -> column T | MIS In Transit (`ships`, `date=max(arrival, today)`) |
| Material of each shipment = your WSA PO + BL# first, then description + supplier + RM Code (`map_in_transit`) | `ppc_arrivals` / `read_ppc_in_transit` | `normalize_in_transit` |

| Front / back glass = two materials (one Balance each) | one sheet per glass (`_glass_split`, below) | per material (always was) |
| Which supplier goes with which customer = Settings ALLOWED VENDORS (first vendor; blank Supplier = any vendor); stock by supplier = MB52 batches (Batch Master) / count by supplier | a material on 2+ sheets (333: SB Back + Subanci Back): each customer on its sheet, stock + PPC arrivals by supplier (`_wsa_shared`, below) | supplier rows (Action List / MRP Daily) + `allocate_lines` (MRP by Line) |
| Settings ALLOWED VENDORS on the projection (`vendor_pools`, user Nov 2026: "WSA output must follow settings ... all output workbooks"): a customer tied to a vendor may use only that vendor + unknown-vendor stock; customers without a rule may use everything; short on day K when a tied group needs more (up to K) than its vendor supplied, or the pooled total is short; lock = pooled - that | every planning sheet: AK 'Vendor rule: not usable' (V subtracts it), AL own balance per vendor, AK3 = the rule; 333 ALERT rows read AK | `_mrp_vendor_pools` -> `x['S']` minus lock (ALERT, Action List, RM Issues, dashboards follow); ALERT Calc 'Projected stock' subtracts the lock change |
| ALERT list = under 21 days of stock, a stock-zero day inside the lead time, or below 21 days within 16 weeks (old Risk Charts rule) (`alert_eval` + `alert_why` + `build_alert_rows`, built ONCE) | Inbound first tab 'ALERT' (`add_alert_inbound`, the same rows) | first tab 'ALERT' (`write_alert_sheet`) |

Nov 2026 check: 280 shipments, 7,056,776 in transit (MRP units) on 20 materials - identical per material in both outputs.

Run: `python update_inbound.py` (Python 3.9+, pandas, openpyxl, lxml). A run takes about 50 s: every run also writes the
aligned WSA file (see `--align-wsa` below; `--no-align` = old behaviour). Step 1 stops `build_new_mrp` right after `align_capture()`
(`_CaptureDone`, ALIGN_MODE only) - it used to write a full MRP_Daily to a temp folder and throw it away (Nov 2026: 69 s -> 50 s, aligned
file byte-identical, outputs cell-identical). Last console line: 'Timing: step 1 ... aligned WSA file ... step 2 ... total' (~10 / 11 / 28 s).

## Key rules (how the numbers should behave)
- **Demand = module plan x quantity per module**, split by customer and line. Not the WSA sheet demand.
  - Module plan: PPC `Plan`, with PPC `Revised Plan` on top (for each date + production line in the Revised Plan, its rows
    replace the `Plan` rows of that date + line). Without a PPC file: `Production Plan` sheet of MRP.xlsx.
    Code: `ppc_plan_sheet()` / `read_revised_plan()`, served as 'Production Plan' through `_RedirWB`.
  - Revised Plan customer names are matched to `Plan` names (typos like 'Clenera Crimsion'); Settings MIS LINK overrides
    accept a near-identical spelling (`link_override()`).
  - Quantity per module: PPC `BOM` (`Qty per Module`) for the customers it lists, else the MIS Standard BOM.
    Customers of the plan with no BOM rows (e.g. Banita, Rowdy Creek, Desri) and MIS customers not in the plan use MIS.
- Match plan customers to MRP customers through `LINKS` (plan name -> MIS customer), never by exact spelling.
- **Stock = your MB52**, replaced by the physical count from the **PPC `Count` sheet ONLY** (rolled forward to today).
  No PPC file in the folder -> no counts at all (WARNING, MB52 for every material); MRP.xlsx / previous MRP_Daily
  'Stock Count' sheets are never read for counts. The count used = the newest date column with a number.
  A 0 in the newest date column is a REAL count of 0 (user confirmed Nov 2026: pine wood pallets PMPKM00000447 / 535 on 05-Nov).
- **OEM** (customers with OEM = Yes on Settings): their plan and the **JW stock at the OEM** are left out of every projection,
  shortage date and risk list.
  **JW STOCK AT THE OEM = in no calculation** (user rule, Nov 2026, final wording: "ignore stock at OEM only"; two earlier readings
  were wrong - do NOT remove / zero OEM data and do NOT change the OEM target or plan logic). The JW stock stays visible as data
  (Action List 'to OEM Stock', Stock Data 'to OEM' lines, No Demand / Movement Log JW, WSA OEM customer columns keep their modules);
  it is not added anywhere: Action List 'Inv. Days at line daily target_with OEM' = (own stock + in transit) / target with OEM
  (single material and substitute group), pack-window qty = own stock only. 'Below 21 Days From' / 'Days Late' (target) still use
  the target WITH OEM (`x['cw']` from `tgt_need`).
  In code: the OEM step in `build_new_mrp` subtracts `st_jw` from `b['stock']` once and sets `b['jw_out']=True`.
  Any code that runs after it must **not** subtract `st_jw` again (check `jw_out`). This was the Nov-2026 Stock Check bug.
  Code that runs BEFORE the OEM step (Check sheet, glass supplier split) uses `own_stock(b)` (stock without JW, either side of the step).
  Nov 2026 audit (user: "make sure OEM stock is not included in all output calculation"): leak test = MIS 'JW WSA to OEM' x10 + 77,777
  -> only Stock Data 'to OEM' lines, No Demand JW column / its 'other RM' list and Movement Log change. Two leaks fixed then: supplier
  'Make' stock read RM Status column L (JW) -> 280 XINYI showed 59,009 (13,959 own + 45,050 JW), 268 XINYI 62,837 (21,717 + 41,120);
  Check 'MIS stock vs WSA sheet balance' compared stock + JW -> 7 false High issues (JB 500, M10R Frame, Potting A/B, 0.26 Ribbon, 280, 268).
  Also: the OEM step now takes out a NEGATIVE JW too (MIS has -8.7e-11 on RMTAP25), Check 'Usage per module' severity uses own stock.
- **MRP.xlsx = Settings only is valid** (Nov 2026: the planner's MRP.xlsx now has Script / Settings / Index / Common Name / Material
  Group / Batch Master, no Production Plan / BOM - those come from the PPC file). `_is_user_mrp` accepts a file with a Settings sheet;
  before, such a file was ignored WITHOUT a warning -> no Settings (OEM customers, targets, substitutes, allowed vendors), no PPC counts,
  no packaging (ALERT 13/8/0/2 instead of 15/19/2/1). Console must say 'Settings read from MRP.xlsx' and 'Stock rolled forward ... counts'.
  Sheet texts (Module Plan / BOM / Stock Count A1, MRP Daily D1, Check, No Demand, How it works, Definitions) name the real sources:
  PPC Plan + Revised Plan / PPC BOM / PPC Count, Settings ALLOWED VENDORS (no 'Customer Supplier' sheet any more).
- **Check 'Usage per module: MIS BOM vs your BOM'** (`bom_rate_checks`) compares PER CUSTOMER (customers sharing one MIS BOM name were
  added up: 356 showed 2 PC instead of 1) and uses an alternative code (PPC BOM '_Alt' line / other roll length, via alias) only when
  the main code is not listed (busbar 47 showed 0.087 = 47 + _Alt 16 + 32). The demand itself never had either problem.
  Nov 2026: 32 rows / 21 High -> 12 rows / 3 High (RTV 0.39-0.405 KG vs MIS 0.351 KG).
- **Substitute groups** (Settings SUBSTITUTES): projected as one pool (group stock + in transit vs group demand).
- **All sheets of MRP_Daily are visible** (user's choice, Nov 2026). The self-check keeps skipping Check / Stock Data /
  No Demand / Substitute Advice (not in shortage order).
- **Daily targets and OEM = Action List ONLY** (user rule, Nov 2026). Columns 'Production Daily Target_with/without OEM',
  'Inv. Days at line daily target...', 'to OEM Stock' stay there. **Every other sheet follows the module plan**:
  `x['pl_need']` = average plan per production day, next 16 weeks (OEM plan left out, substitute group pooled);
  `x['pl_days']` = (own stock + in transit) / pl_need. Used by the ALERT (Need/day, reorder level),
  Cover Heatmap, risk class 'Under 21 d', risk order, Substitute Advice, the daily report, No Demand, MRP by Line
  ('This line: daily plan need'; the Action List target / OEM columns are not copied there).
  Action List shows BOTH scenarios (user rule, Nov 2026 - do not ask again):
  - daily TARGET: 'Below 21 Days From' / 'Days Late' use the target need (`x['cw']`, with OEM target), next to the target columns;
  - MODULE PLAN (gold band): 'Daily plan need (module plan)', 'Inv. Days at module plan (stock + in transit)',
    'Below 21 Days From (module plan)', 'Days Late (module plan)' (`x['cw_p']` = ALERT 'Below 21 days from' / 'Days late'). Copied to MRP by Line.
  - Status follows the plan. Stock-zero periods on the Action List = 'Short from → to (with in transit)' only (all periods,
    ALERT label, lead-time slack). 'Short From (Stock = 0)' was removed from the Action List (duplicate, user OK Nov 2026);
    it stays on RM Issues.
- **MRP Daily / Weekly layout** (Nov 2026): 'Firm receipt (Open PO, no vessel)' and 'Days of cover (stock + in transit + Open PO)'
  rows exist only for materials with an Open PO (rowmap `ro` / `rj` = None otherwise - guard every use). Supplier column shows the
  single real supplier (e.g. FLAT) instead of 'All suppliers' (`b['_sup1']`; row keys unchanged). Supplier blocks only when 2+ suppliers
  supply the material (`multi_sup`). Days of cover = projected stock at day end / average plan usage of the next COVER_AVG_DAYS (21).
- **No 'Action' column on any sheet** (user, Nov 2026: "Column Action in any sheet no help to me"): removed from ALERT (both outputs),
  RM Issues ('Action / Mitigation' + hidden '_auto'; table A:U) and both Check tables. Do not add action-text columns again.
- **Action List = RM Issues look** (user, Nov 2026: "Action List need to redesign, I want it look like sheet RM Issues"; the user skipped
  the layout question, so the recommended option was built): title 20 pt navy, grey note, header row 4 light grey (F4F6F8, medium
  line below), thin row lines, no banner row 3, table without stripes, Calibri 10, row height 30+; 'own' in `unify_workbook_style`.
  The 8 stock-source columns (Source, Source code, Qty source / MRP unit, UOMs, Factor, Storage Loc.) were dropped (they are on Stock
  Data); 'Priority' added after UOM = the RM Issues rule (High = stock zero inside the lead time, Medium = short later, Low = under
  21 d, OK, No demand; coloured like RM Issues). Every other column, its value and order unchanged (checked: 0 differences on the
  27 common columns + MRP by Line). Nov 2026: 14 High / 24 Medium (incl. 4 glass supplier rows) / 7 OK / 2 No demand.
- **'ALERT (formulas)' + 'ALERT Calc'** (user, Nov 2026: "a new same sheet with excel formulas so that I can check your work"; both
  outputs, right after the ALERT; `write_alert_formulas`). Same rows / column positions as the ALERT (Q = stock-zero days in the plan
  (count), X = in transit arriving after the plan end instead of the two text columns). Yellow inputs: Stock today, On hold, Lead time,
  Overdue. Plan and Arrivals per day = SUMIFS of 'ALERT Plan lines' / 'ALERT Arrival lines' (user, Nov 2026: "I want to see formulas ...
  Plan (demand), Arrivals (in transit)"; U In transit = SUMIFS up to the plan end). `alert_lines()` builds them: plan line = customer x
  production line x day, Demand = Modules x Qty per module x Factor (`b['plan_lines']`, kept in `rm_plan_from_mis` and both packaging
  builders; OEM out via `OEM_IS`); arrival line = one MIS shipment (SHIPS_, counted on MAX(arrival, today)) + Adhoc / substitute per day
  (group: minus members' own stock = `_adh`); group = members x fx; 333 per sheet = the customer / PO subset that adds up to the sheet
  (`_subset_match`). Every row is verified against `r['P']` / `r['Tp']`; a row that does not add up gets per-day projection lines +
  console NOTE. Nov 2026: 15,981 plan lines, 275 arrival lines, all 37 rows add up. Speed: the line sheets get no per-cell font (shared
  fill object) and are skipped by the `unify_workbook_style` passes (run 103 s -> 69 s; checker finding). ALERT Calc rebuilds Projected stock (= prev + arrivals - plan), Stock zero
  (1/0), Below 21 days (1/0, day < 112); every other column is a formula (AVERAGEIFS / SUMIFS / MATCH / `_xlfn.MINIFS` - the prefix is
  needed or Excel shows #NAME?). Y 'Check vs ALERT' (Z = what differs) compares each calculated column with the ALERT row. No cached
  values (fullCalcOnLoad). Proof Nov 2026: LibreOffice recalculation = 37 / 37 OK in both files; planted errors (stock +50,000, lead
  time +20) -> DIFF with the right columns named.
- **Stock Check: GR + current month** (user, Nov 2026: "I update in transit and stock count daily ... Received since count, I see no
  number. I want opening and closing stock of the current month, total GR qty and used qty current month"). MIS 'GR. Date' is saved as an
  Excel NUMBER (46330 = 04-Nov) -> `_any_date`; `read_gr_lines` = PPC GRN on GR. Date per traced code (alias / factor) -> `GR_LINES`;
  'Received since count' = `gr_between(code, count date, today)` (was Final AtA; shown as 0, not blank, for counted materials - with
  daily counts it is mostly 0: a GR before the count is inside the count). `read_count_history` -> `COUNT_HIST` (base unit, RM) /
  `COUNT_HIST_BOM` (Qty_BOM, CS / PM = what the MRP uses): stock on every PPC Count date column (several rows: each row's latest count
  on / before the day, every row needed). The CURRENT MONTH band (Stock Check columns P..V: opening / GR / used / closing /
  projected month end) was REMOVED (user, Nov 2026: "I dont want this") - do not add it back.
- **Stock columns = ONE rule (`stock_src` / `sup_split` / `alert_stock_src`, user, Nov 2026)**: 'Estimated stock today' = PPC count
  rolled forward (count − usage + GR since the count), not counted = SAP (MB52) = the stock every shortage date uses (JW out).
  ALERT column G renamed 'Stock today' -> 'Estimated stock today' (value unchanged) + H 'SAP stock (MB52)', I 'Physical count' (count
  as counted, 'no counting'); every later ALERT column moved 2 right (both outputs; `_ins_cols` at the end of `write_alert_sheet`).
  'ALERT (formulas)' keeps its own positions; its 'Check vs ALERT' points at the ALERT's own letters (see the INVENTORY / QUALITY rule below). Group row = members
  x fx; 333 per WSA sheet = the suppliers whose estimated stock adds up to the sheet's stock (`_subset_match`, r['sups']; none = '-').
  Material Status 'Stock SAP' / 'Stock Count' -> 'SAP stock (MB52)' / 'Physical count' = the same values (Other Materials sheet
  of Daily_Material_Status too). Stock Check: column B 'Material' -> 'Item name' (= ALERT, MIS Common Name) + C 'Category'
  (`item_category`), every later column +1; 'Estimated stock today' filled for every row (not counted = SAP); a material with stock
  from 2+ suppliers = one row PER SUPPLIER ('<name> - FLAT'): SAP = MB52 batches (`b['sup_sap']`, saved in `attach_wsa`), count =
  PPC count by Supplier Code (`b['sup_cnt']`, `apply_count_by_supplier`), estimate = the MRP Daily supplier row, usage / GR by count
  share, rest = 'Unassigned'; 333 supplier rows carry the ALERT label of their WSA sheet.
- **No INVENTORY / QUALITY on the ALERT or Material Status** (user, Nov 2026: "Alert sheet, delete quality and inventory. Do same in
  material status"): Reorder level, Status, On hold columns + the ALERT 'INVENTORY · QUALITY' box + the Material Status INVENTORY /
  QUALITY cards removed (`_del_cols` at the end of `write_alert_sheet` / `add_stock_mgmt_dashboard`; Other Materials too). ALERT
  now 23 columns: A-F, STORE G-J (Estimated stock today, SAP stock (MB52), Physical count, Days of stock), PRODUCTION K-L, TRANSIT
  LEAD TIME M-P, BELOW 21 DAYS Q-S, WAREHOUSE T-W. The fields (rop / status / hold) are still computed (`alert_stock_fields`);
  'ALERT (formulas)' keeps its own columns (incl. reorder / status / on hold) and compares only the columns the ALERT still has.
- **Daily_Material_Status sheet 2 = 'Other Materials'** (was 'Indirect Material'; user, Nov 2026): PPC Count materials not in the BOM.
  Need / day = PPC Count column 'Material Req Qty per Day' (rows of one material added up; `make_daily_material_status`); a counted one with a need = count rolled forward (- need x days since count + GR, checker finding) -> Days of
  supply = stock / need, Days incl. in transit = (stock + MIS in transit) / need, Short from = first day stock + MIS arrivals (late =
  today) - need x days < 0 (180 days), Shortage risk 'Short dd-mmm' / 'Under 21 d' / OK / No demand (no number). Nov 2026 file: column K
  is an XLOOKUP into the PPC 'Usage' table and has numbers only for BOM materials -> 0 of 92 rows have a need (tested with planted
  values: 13,727 / 200 -> Short 16-Jan, 277,050 / 20,000 -> Short 22-Nov). Not the Inbound 'Other materials' sheet (BOM materials
  without a WSA sheet) - different file.
- **UOM on Material Status + Daily_Material_Status** (user, Nov 2026: "Add UOM into daily material status"): column F 'UOM' after Category
  (`_add_uom_col`, last step of `add_stock_mgmt_dashboard` / Other Materials; band A1:F1, every later column +1). Material Status /
  Direct Material = the MRP unit (b['uom'] / ALERT row uom); Other Materials = PPC Count 'Base Unit of Measure' ('-' when blank).
  Daily_Material_Status Direct Material has NO STORE / WAREHOUSE / PRODUCTION cards (user, Nov 2026: "No need from daily material
  status"; removed after `_copy_sheet_to`, notes moved up); the MRP_Daily Material Status sheet keeps them.
- **Daily_Material_Status = 2 sheets** (user, Nov 2026): 'Materials Status' = Direct Material (MRP_Daily Material Status copy, no cards)
  + a navy band row 'OTHER MATERIALS ...' + the Other Materials rows (Sr. continues), then the summary boxes / notes; columns F 'Item
  Category' = PPC Count column I 'Item' (by code), H 'Consumable Unit (Production Confirmed)' = PPC PM_CS_Usage (by Code; '-' = #N/A /
  blank / '-'); 'ALERT' = copy of the MRP_Daily ALERT WITHOUT the 4 class cards (`_drop_card_rows`; definitions line moved up). Built
  as 2 temp sheets then merged (`make_daily_material_status`). Nov 2026 (new PPC file): 45 BOM + 92 other rows.
  Both sheets end with 'Usage per module by customer (source: your BOM / MIS Standard BOM)' (= Stock Check / old Action List
  `x['usage_src']`, by code; other materials / none = '-'; user, Nov 2026).
  + 'Cover Heatmap' (copy of MRP_Daily's) + 'RM Material Dashboard' (= the daily WSA 'Dashboard' sheet; `_wsa_dashboard_to`: every
  VLOOKUP formula worked out in Python on the aligned planning sheets' values, date row = today + 1 per column - the WSA file's own cached
  Dashboard results are stale (Excel's last calculation) and the Back rows have none; read-only load with a light style table (fonts /
  fills / borders, no 36,000 named styles; the global speed patch drops them); widths / pane / zoom / the 4 colour rules from the sheet
  XML. Proof Nov 2026: 3,033 cells = LibreOffice's recalculation of the aligned file's Dashboard, 0 differences; run +6 s).
- **'Cover Heatmap' renamed 'BOM Material Dashboard'** (user, Nov 2026; `HEAT_NAME`, both MRP_Daily and Daily_Material_Status). The
  Daily_Material_Status 'RM Material Dashboard' has column C 'Category' (was a thin spacer) = `item_category` of the WSA sheet's codes
  (`WSA_REFS` groups, only the sheet's MRP material(s) in AL_ORDER_ - not old / alias codes; several = joined ' / ', one per name in any
  case; no plan sheet = '-').
- **'RM Material Dashboard - MIS'** (Daily_Material_Status, user Nov 2026: "I want same story ... in new sheet"; `DASH_MIS`): the same
  WSA Dashboard, daily columns read column V 'Plant + Port Days (PPC In Transit)' = MIS arrival dates (`_wsa_dashboard_to(use_v=True)`:
  A=19 -> column 22) instead of S (the planner's typed Scheduled Delivery); D (Plant Days today) unchanged; own risk order. The old RM
  Material Dashboard (column S) stays. Every daily cell = V of its OWN date (the planner row Frame G12R 35HS reads R 2 days shifted - not
  copied); column A = 22. Nov 2026: first value below 1 = the ALERT first stock-zero day on every planned row except the EPE group (known).
  BOM Material Dashboard (weekly) divides by the AVERAGE need per production day. **'BOM Material Dashboard - Daily'** (user, Nov 2026;
  `DAILY_NAME`, `_bom_daily_sheet` built LAST (after the ALERT, before `unify_workbook_style`; HEAT_ROWS) so no fixed-index tab placement
  moves; MRP_Daily + Daily_Material_Status copy; Definitions row): same rows / groups / order; one column per day (141 today) =
  stock available that morning (S_k + P_k = stock + MIS arrivals to day k - plan of the days before) / THAT day's demand ('na' = none) =
  the WSA column V rule; D = stock today / today's demand; the 4 WSA Dashboard colour rules (material rows only - a blank band cell would
  turn red). Nov 2026: first value < 1 = ALERT first stock-zero day on 43 / 44 rows (EPE 480 (own) vs pooled group = known).
- **Dashboards Nov 2026 (user: "Daily Material Status, all dashboard must sort from highest to lowest risk. Hide sheet RM Material
  Dashboard. Rename BOM Material Dashboard to BOM Material Dashboard - Weekly. BOM Material Dashboard - Daily format and design same as
  weekly. Hide sheet alert")**: `HEAT_NAME` = 'BOM Material Dashboard - Weekly' (both outputs; the old name is deleted on rebuild).
  ONE risk order for every dashboard (`bom_risk_key` on `bom_day_cover`; RM Material Dashboards `_risk`): first day the daily cover
  (stock that morning / that day's demand) goes below 1 (= stock-zero day; checker finding: below 0 missed 1-day stock-outs such as
  Frame Subanci / Busbar 5mm on 12-Nov), then below 5, then the lowest cover today; no demand last; inside the RM / PM / CS
  groups (weekly + daily share the order; the Alert column is no longer the sort key). Daily sheet = the weekly design (dark heatmap,
  title + note row 2, header row 4: Code / Material / Category / Alert / Today's Morning stock, days from F, static fills with the weekly
  bins - short = below 1 day, na grey; legend + source line; frozen at F5; skipped by the `unify_workbook_style` code-cell pass).
  Daily_Material_Status: 'ALERT' and 'RM Material Dashboard' are HIDDEN (still built; MRP_Daily ALERT stays visible - all sheets visible).
  Check Nov 2026: every weekly / daily row keeps its values (45 / 44 rows), only the order changed; all other sheets cell-identical.
- **Daily_Material_Status layout (user, Nov 2026: "sheet Materials Status, delete row 1, then hide this sheet; sheet ALERT delete row 10
  then move row 1 to 9 to bottom")**: done at the END of `make_daily_material_status` with `_remap_rows` (cells, styles, merges, heights,
  freeze panes, autofilter, print titles follow): Materials Status = no group band (header row 1, data from row 2, frozen C2) and HIDDEN;
  ALERT = no group band (header row 1, data from row 2, frozen E2, filter A1:X..), the summary boxes below the table after one blank row.
  Visible tabs: BOM Material Dashboard - Weekly (opens first), - Daily, RM Material Dashboard - MIS. MRP_Daily / Inbound ALERT unchanged
  (presentation of the daily file only).
- **RM dashboards in the BOM design (user, Nov 2026: "apply color rule mentioned in row 55 from bom material dashboard to 2 sheets rm
  material dashboard ... rm dashboard design style same as bom dashboard. Then hide bom material dashboard daily")**: `_dark_rm_dashboard`
  rewrites 'RM Material Dashboard' + DASH_MIS (Daily_Material_Status only, after `_wsa_dashboard_to` - `sort_wsa_dashboard` / DASH_MAP
  untouched) as the dark heatmap: title + note, header row 4 (A lookup column, B planning sheet, C Category, D Today's Morning stock,
  days from F), data from row 5 (same rows / order / values, 0-0.99 with 1 decimal), static fills by `DAY_BINS` (= the BOM Daily legend
  row: short < 1, 1-7, 7-14, 14-21, 21-35, 35-60, 60+, na grey; the weekly legend has short < 0 = week-end stock), No Plan = grey chip,
  legend + source line, frozen F5; the 4 WSA colour rules dropped. Daily_Material_Status visible tabs now: BOM Material Dashboard -
  Weekly + RM Material Dashboard - MIS (BOM Daily hidden there; MRP_Daily keeps it visible).
- **Code + Item name on every dashboard (user, Nov 2026: "Daily material status output, I want code and items name in all dashboard")**:
  BOM Weekly column A = Code (was a spacer) + header 'Item name' over B (both outputs, same builder; skipped by the code-cell style pass
  like the daily); BOM Daily header 'Item name'; RM dashboards (`_dark_rm_dashboard`): A Code(s), B Item name(s) (MIS Common Name =
  `b['common']`) of the WSA sheet's MRP material(s) (`DASH_CODES` from `_wsa_dashboard_to`, same codes as Category), C Planning sheet
  (the WSA Dashboard label), D Category, E Today's Morning stock, days from F; the lookup column index is left out; No Plan rows '-'.
- **Weekly = the daily number on each week's date (user, Nov 2026: "bom dashboard dec 01 should same as dec 01 in rm dashboard"; the
  user declined the option question, so the recommended option was built)**: `add_cover_heatmap` weeks(r) = `bom_day_cover(r)[7w]` (cover
  that morning / that day's demand on d0 + 7w; 'na' = no demand that day) instead of week-end stock / average need; colours = DAY_BINS
  (short < 1), legend DAY_LABS + na, below-1 values floor to 1 decimal; both outputs. Check Nov 2026: weekly = BOM Daily on the week date
  (value + colour) on every row with demand (the 2 no-demand rows stay blank on the weekly, 'na' on the daily); 01-Dec weekly = RM
  Material Dashboard - MIS on all 23 planned rows. Full item names (no [:34] cut) on weekly + daily; weekly band rows filled from column A.
- **Weekly dashboard E..J = Material Status columns (user, Nov 2026: "Add these from material status sheet into bom material dashboard:
  In transit, Next arrival, Need / day, Short from, Days of supply, Days incl. in transit")**: `HEAT_MS_COLS`; `add_cover_heatmap` leaves
  E..J (weeks now from K) and records `HEAT_MS_ROWS`; `_heat_ms_fill` (after Material Status is built, before the daily sheet / style /
  Daily_Material_Status copy) copies value + number format from the Material Status row by code (333 by its '(WSA sheet: ...)'). EPE
  (own) rows (user answer Nov 2026: option 2 = own values): In transit / Next arrival from Material Status (already per code), Need / day
  = the member's own need, Short from = its own first stock-zero day (= the Daily dashboard's first value below 1: EPE 480 10-Nov, EPE
  420 10-Dec), Days = own stock (+ in transit) / own need (`own_ms`); Material Status itself keeps the pooled group values. Short from
  dates in bold red text. Both outputs.
- **Cover Heatmap grouped RM -> PM -> CS** (user, Nov 2026) by the PPC Count 'Item' (`count_items`; user: "refer to item category from
  count ppc"): Packaging -> PM, Consumable -> CS, any other item -> RM; not on Count -> code prefix (RMLAB = PM, RM / PM / else CS).
  Nov 2026: RM 26 / PM 19 (incl. RMLAB labels, Kraft labels CST 2254 / 2255, and Perforated Tape RMTAP25 - its Count Item is
  'Packaging') / CS 0 (empty group: no band). Column C 'Category' = `item_category` (= Material Status). A band row per group, risk order
  inside each group; `audit_workbook` FREE = ALERT + Cover Heatmap. Same rule for MRP Daily / Weekly column C (`_short_lbl`, `COUNT_ITEM`):
  a 'Raw material' / 'CS - Consumable' label of a Count-'Packaging' (or RMLAB) code -> 'PM - Packaging'; other labels unchanged.
  RMLAB labels also: MRP Daily / Weekly column C 'PM - Packaging', Inbound 'Other materials' Type 'Packaging / consumable' (labels only -
  `is_rm`, lead times and RM Issues unchanged; RM Issues membership = open question to the user). Daily_Material_Status Materials Status:
  the other-material rows sorted by risk too (Short dd-mmm earliest first, Under 21 d, OK, No demand; checker finding).
- **User answers Nov 2026: tape = PM, barcode / back label = PM, sort the WSA Dashboard = yes.** `is_pm_like(code)` = PPC Count Item
  'Packaging' (a code not on Count takes its alias code's Item: Cell Fixation tape RMTTT010 <- 018) or RMLAB -> Cover Heatmap PM group
  (RM 25 / PM 20), MRP Daily / Weekly label, and OFF RM Issues (24 -> 20: RMTTT010, RMLAB121 / 097 / 098). `is_rm` / lead times unchanged.
  `sort_wsa_dashboard()` (after `align_wsa`, step 1): the aligned file's own Dashboard rows re-ordered at XML level in the RM Material
  Dashboard order (`DASH_MAP` from `_wsa_dashboard_to`); each formula's own-row refs ($A5) follow the row (text in quotes untouched:
  'Glass G12R ...' is not a cell), unquoted 'Dashboard!B5' refs on other sheets (hidden Summary Dashboard, 15 refs) are moved too; read-back
  check before replacing. Proof: LibreOffice recalculation - every row keeps its values, Summary Dashboard labels unchanged.
- **RM Material Dashboard sorted highest risk first** (user, Nov 2026: "All dashboard, sort from highest risk to lowest"): by the first
  date its cover (Plant + Port Days) goes below 1 (stock-zero day; was 0 until Nov 2026), then below 5, then the lowest Today's Morning stock; No Plan rows last; values unchanged.
  (ALERT / Materials Status / Cover Heatmap were already in risk order; the WSA file's own Dashboard is not re-sorted - its formulas
  point at fixed rows.)
- **Packaging stays on the ALERT** (user answer Nov 2026: yes) - the 16-week rule lists RM AND packaging (PM).
- **Stock Check 'Shortage risk' = the ALERT label** (user rule, Nov 2026: "ALERT and Stock Check use the same label"; before, Stock
  Check used `risk_info` (days incl. in transit) and showed 'OK' for 280 / 268 / InnerBox 620 that the ALERT lists as UNDER 21 DAYS /
  UNDER 21 DAYS LATER). `alert_label(code)` reads `ALERT_ROWS` (class + first stock-zero day; class 3 + 'from' below-21 date; 333 = one
  line per WSA sheet; substitute member = group row via `grp_main`); not on the ALERT = OK / No demand / No need. Chip colours = ALERT.
  Rows sorted in the ALERT order (first ALERT row of the code / its group), then the rest in the old order. RM Issues still uses `risk_info`.
- **Action List sheet REMOVED from MRP_Daily** (user, Nov 2026: "Delete Action List, no help me at all"). It is still built in memory
  (other sheets / AL_ORDER / AL_VAL use it) and deleted just before the final save; `audit_workbook` skips when it is missing.
  Its 4 'who uses it' columns moved to Stock Check P..S: Lines (share of plan volume), Project Name (share of plan),
  Usage per module by customer (your BOM / MIS Standard BOM), Module (MIS BOM, share of plan).
  Glass PO 6000005131: both GR lines (HTAR, RM Code 312) are traced to 312 (36,000; same tracing as in transit, Check High).
- **ALERT + Stock Management = ONE sheet 'ALERT'** (user, Nov 2026: "combine alert and stock management into 1 sheet, style design stock
  management"; both outputs, first tab). The Stock Management sheet was removed then, and BROUGHT BACK (user, Nov 2026: "bring back
  stock management and include all materials not just raw materials"): `add_stock_mgmt_dashboard` restored from 0a3086c, rows = EVERY
  material in AL_ORDER (29 RM + 13 PM + 1 CS, also No demand; chip added), same columns / rules as before (the 23 old rows identical),
  sheet right after the ALERT; the ALERT keeps its Stock Management columns.
  Rows = the ALERT rows (every old Stock Management row was on it: EPE 420 = the EPE 480 + 420 group row, now pooled). 24 columns in
  Stock Management groups (layout today: see 'No INVENTORY / QUALITY' above - 23 columns, A-F, STORE G-J, PRODUCTION K-L, LEAD TIME M-P,
  BELOW 21 DAYS Q-S, WAREHOUSE T-W In transit, Next arrival, Overdue in transit, Arrival at plant). New fields from `alert_stock_fields` (in `build_alert_rows`, so the
  Inbound copy has them): reorder = need x (lt + SAFETY_DAYS), status Out of Stock / Below reorder / Above reorder / 'Short dd-mmm',
  on hold = MIS MB52 blocked + QI (`mb52_hold`; group = members in the main unit; 333 rows: on the sheet row(s) holding its stock, by stock share - QI 10,800 is KIBING = SB Back; checker finding Nov 2026), in transit = the
  row's tr_tot, overdue = MIS ships before today (333 per sheet: "in today's arrivals"; the summary counter counts materials incl. 333). Nov 2026: 37 rows, the 22 other old Stock
  Management materials identical (stock / in transit / overdue / reorder / status / hold). Below the table: ALERT summary, first
  stock-zero days, INVENTORY · QUALITY, why below 21 days boxes + 4 class cards + definitions line.
- Material status: CRITICAL / AT RISK / WATCH / OK vs each material's own target days (`material_status()`); line status uses COVER_* (`line_status()`).
- **ALERT sheet = first tab of MRP_Daily AND of the Inbound file, the same rows** (user request Nov 2026: "list of less than 21 days
  inventory, how many days stock zero within lead time and zero from-to ... many sheets have this info but too complicated - one sheet,
  simple and strong alert"). `build_alert_rows(A,d0,nd)` (end of `build_new_mrp`, stored in `ALERT_ROWS` / `ALERT_D0`) -> `alert_eval`:
  days of stock = stock today / `pl_need`; days incl. in transit = `pl_days`; stock-zero day = S < 0 (the Action List projection: stock +
  MIS in transit arrived - plan; the day's need is not covered; on a WSA sheet = Plant + Port Days (PPC In Transit) below 1 - its red
  cell (< 0) comes a day later); zero days inside the lead time = S < 0 days from today to today + `lt_used` - 1; listed when days of
  stock < REPORT_DAYS (21) or zero days inside the lead time > 0. Classes: STOCK-OUT IN LEAD TIME (red, EXPEDITE / alternative),
  STOCK-OUT LATER (orange, ORDER NOW / ORDER BY first zero day - lead time, RELEASE_WINDOW like `risk_priority`), UNDER 21 DAYS
  (yellow, in transit covers), UNDER 21 DAYS LATER (blue: 21+ days today, below 21 within 16 weeks, no zero day).
  **Risk Charts merged into ALERT** (user, Nov 2026: "Alert and risk chart can be combined?"): the Risk Charts sheet is gone; its
  columns are on the ALERT (both outputs): 'Below 21 days from', 'Why (cause)' (LATE ARRIVAL / LATE + NOT ENOUGH / NO SUPPLY),
  'Days late', 'Arrival at plant (Final AtA)' (`x['ata_txt']`; 333 per sheet: its PPC arrivals), (`alert_why` = the `cover_watch` rule; 333 per sheet from the sheet's S / T). Listed also when below 21 days within 16 weeks, for
  RM AND packaging (same rule) - Nov 2026: 15 red / 19 orange / 2 yellow / 1 blue (was 15 / 7 / 2; +5 RM from Risk Charts, +8 PM).
  **ALERT look = Stock Management** (user, Nov 2026: "I want alert sheet is designed like stock management"; `write_alert_sheet`, both
  outputs): row 1 = group band (ALERT · date / STOCK TODAY / LEAD TIME · stock-out / BELOW 21 DAYS · why / SUPPLY · in transit), row 2 =
  header coloured by group, data from row 3 with 'Sr.' in column A (code = column C), light tinted group columns, soft chips (Stock
  Management colours), frozen at E3; below the table: 3 summary boxes (ALERT summary / first stock-zero days / why below 21 days), 4
  counter cards (one per class) and the definitions line. Rows, values and order unchanged (checked Nov 2026: 23 rows identical).
  **ALERT Item name + Category** (user, Nov 2026: "Alert sheet, I want item name and category from stock management"): column D
  'Item name' (= Stock Management item name, MIS Common Name; group / WSA sheet notes kept), new column E 'Category' -> every later
  column moved one right (19 columns then; 24 since the Stock Management merge). `item_category(m)` = ONE rule for Stock Management and the ALERT: Material Group
  column E > MIS Standard BOM Understanding Name (`MIS_CAT`, alias) > MRP.xlsx Material Group column A (`MAT_GROUP_A`: Packaging /
  Label; the planner's sheet has no column E today) > '-'. Nov 2026: 37 rows, no '-'; Stock Management Barcode / Back Label '-' -> 'Label'.
  Stock Management shows the full name / category (no [:30] / [:16] cut). ALERT row notes kept: substitute group = all members, 333 + WSA sheet.
  **Cover Heatmap** = the SAME rows as the ALERT (`add_cover_heatmap(path, ALERT_ROWS, ...)`, column C = alert class), sorted high
  to low risk (user request Nov 2026): class, then most weeks out of stock, then most weeks under 21 days, then first zero day. Columns: Alert, Code, Material, UOM, Stock today, Need / day, Days of stock, Days incl. in transit, Lead
  time, Stock-zero days inside the lead time, First stock-zero day, Stock zero from -> to (every period; '⚠' = starts inside the lead
  time), Below 21 days from, Why (cause), Days late, Next arrival, Arrival at plant (Final AtA) (17, no Action since Nov 2026); four coloured counters + one definitions line on top. Substitute group = one row (`grp_label`, pooled;
  followers skipped) on the ALERT - but on the Cover Heatmap the group is split per material (user, Nov 2026: "EPE 480 + EPE 420
  (group), i want separate"): `r['mem']` (build_alert_rows) = each member's own stock, plan, transit + adhoc (`_adh` rule), own
  projection + `alert_eval` class ('Not on ALERT' chip when the member alone would not be listed); label '<name> (own)'. Checked Nov
  2026: weekly cover = the EPE 480 GSM / EPE 420 GSM WSA sheets (Balance + PPC arrivals - Demand) on 10 / 10 weeks. A material on 2+ WSA sheets (333) = one row per sheet from the aligned sheet's own numbers (`WSA_REFS[0]` groups:
  today's balance, Demand, PPC arrivals). MRP_Daily: `unify_workbook_style` treats ALERT as 'own' (red tab kept; its code-cell pass
  centres code cells - the sheet centres all its rows the same way); `audit_workbook` FREE = order not checked (its own order: class,
  first zero day), the 'labelled safe' rule still is. Nov 2026: 14 red / 6 orange / 2 yellow; every WSA sheet recomputed (today's
  balance + PPC arrivals - Demand) gives the same stock-zero starts as the ALERT (16 of 17); the exception is the substitute group EPE
  480 + EPE 420 (pooled on the ALERT: 19-Jan; EPE 420 sheet alone: 09-Dec, EPE 480 alone: never). 333: Subanci Back red (FLAT, zero
  12-Nov), SB Back orange (KIBING, zero 20-Jan) - the pooled Action List row says 14-Jan.

- **WSA alignment = DEFAULT on every run** (Nov 2026: the user runs plain `python update_inbound.py` and expected the WSA file;
  `--no-align` skips it, `--align-wsa` is still accepted). Step 1 = calculation run (temp COPY of the WSA file, outputs to a temp
  folder, console muted by `_QUIET` except errors); `align_capture()` (in `build_new_mrp` right after the OEM step) keeps the
  numbers; `align_wsa()` saves the result OVER the user's WSA file (`<stem>.xlsx`, stem without any '(aligned <date>)'; user, Nov 2026:
  one WSA file): first `_wsa_backup` copies the file as it was (and the file the run read, if another) to '<folder>/old WSA files/<name>
  (before <date> <time>).xlsx' (newest 10 kept), then `_wsa_tidy` moves leftover '<stem> (aligned <date>).xlsx' files of older versions
  into 'old WSA files' (never deleted). Step 2 = normal run on the updated file (subprocess, `--no-align`); the #REF! check compares with
  the temp copy of the file as it was. An alignment error is reported and step 2 runs on the user's file (unchanged).
  The console says 'WSA FILE UPDATED: <path> (... backed up in 'old WSA files': ...)'. Proof Nov 2026: run 1 on the planner's file =
  the earlier aligned file part by part; run 2 on the updated file = identical; leftover aligned file moved; backup = the original.
- **Redesigned planning sheets** (Nov 2026 user request: "delete all rows from Oct backward, same format, same header position"):
  every visible planning block -> one sheet in ONE layout (`WSA_FIX1` + customer slots + `WSA_FIX2`): rows 1-6 info (A2 stock,
  A3 'Supplier & Material'/C3, A4 'Unit of Measurement'/C4, A5 codes from C5, A6 totals from today), header row `WSA_HDR`=7,
  one row per day from the 1st of the current month (`keep_from`) to the last date (31-Mar-27). A..E Date, Day, Balance, Planned,
  Demand; F..M = `WSA_SLOTS`=8 customer columns, the same labels in the same order on every sheet (plan labels via the WSA 'Lookup'
  sheet, own lines by modules, then OEM, then labels with numbers on kept past rows); N..S Actual modules, Actual, MW, Actual MW,
  Plant Days, Plant + Port Days (S = col 19); T/U/V = PPC In Transit arrivals / Supplier Name / Plant + Port Days (see below);
  W Delivered, X Scheduled Delivery .. AF Vessel ETA, AG Status, AH Invoice, AI Reference, AJ Remarks 2 (all by key: `XL(k)`).
  **Front / back glass on their own sheets** (user rule, Nov 2026: "balance is updated wrong ... BOM uses both codes ... I want front
  glass and back glass separated in different sheets"; the old combined sheet showed ONE Balance = front + back stock, e.g. Glass Topcon
  39,630 = 280 front 17,913 + 268 back 21,717): `_glass_split()` (start of `align_wsa`) turns a block with front + back glass (side by
  `_glass_side` = MIS Common Name Front / Back / HTAR / Printed) into '<sheet> Front' + '<sheet> Back' (Glass Topcon -> Glass Topcon
  Front 280 + Back 268; Glass G12R Topcon SB -> SB Front 312 + SB Back 333). The source sheet is RENAMED to the front name (all
  references in other sheets / defined names follow, `_ref_fix`; the Dashboard row reads the front sheet; back sheets get their own
  Dashboard row below the table - user answer Nov 2026 'Yes': a copy of the front row reading the Back sheet, label '<label> Back',
  autofilter / conditional format / hidden _xlnm._FilterDatabase name extended, no cached value (Excel calculates); a rerun finds it and adds nothing). A one-glass block whose BL# lines show the partner glass too (Glass G12R Topcon Subanci = 356 + 333) is split the same way:
  'Glass G12R Topcon Subanci Front' (356, the renamed sheet) + 'Glass G12R Topcon Subanci Back' (333) - each customer project keeps its
  own sheets (user, Nov 2026: "you have the demand of each customer; follow the MRP daily settings and definitions to calculate the WSA
  daily, so you know which supplier goes with which customer"; the earlier version MERGED Subanci's 333 into SB Back = one 333 pool -
  wrong: Subanci may use only FLAT, SB Energy only KIBING). An aligned file of that older version is detected by its A1 notes ('its back
  glass RM...' / 'includes the back glass of sheet') -> WARNING: run on the planner's own WSA file. Every quantity
  of the planner's rows is split: Scheduled / Delivered by the MIS In Transit lines of the row's BL# (`_it_bl_index` - traced codes,
  Final Qty.), else 1 front : 1 back; Balance / Demand / Actual / PPC arrivals of kept rows 1:1; text only on rows with that glass.
  Archive records of a block split in this run are split the same way (`info['arch']`). New sheets carry '(front|back glass only ...)'
  in A1 = marker: the next run keeps it and never splits again (same-day rerun on the output = byte-identical file). Nov 2026: 56 rows
  split by BL#, 224 (mostly deleted pre-Nov rows) 1:1; totals conserved; MRP_Daily Check lost 16 'Same shipment, different material'
  (Subanci back glass) warnings. Balances: Front 280 17,913 / Back 268 21,717 / SB Front 312 2,764 / SB Back 333 KIBING 2,931 /
  Subanci Front 356 11,628 / Subanci Back 333 FLAT 5,168.
  **One material on 2+ sheets** (`_wsa_shared`, start of step 1 in `align_wsa`; today only back glass 333): customers = the labels with
  numbers in the sheet's customer columns on rows from the 1st of the month (the planner's sheets keep old customers on 2025 rows - SB
  had 'Subanci', Subanci had 'SB Energy G12R'), else the sheet whose customers share the first allowed vendor, else the first sheet;
  demand not split by customer -> first sheet. Stock by supplier = `align_capture` 'sup' (MB52 batches via Batch Master, count by
  supplier if counted, as the MRP_Daily supplier rows): a supplier goes to the sheet whose customers have it as FIRST allowed vendor
  (2+: by demand); other / 'Unassigned' stock -> sheets with any-vendor customers (Encap), else the first sheet. PPC arrivals: the sheet
  where the planner typed the PO + BL# if its customers may use that supplier, else by supplier. Row 2 (E2) says the split; console line
  'RMGLS00000333 on 2 sheets ...'. Nov 2026: SB Back = SB Energy (KIBING) + Encap (any) - stock KIBING 2,931, arrivals 567,139 (KIBING
  430,189 + FLAT HONG KONG PO 6000005517 136,950 = Encap's PO on the planner's SB sheet); Subanci Back = Subanci (FLAT) - stock FLAT
  5,168, arrivals 295,320 (FLAT VIETNAM PO 6000005271). MRP_Daily 333 supplier rows: FLAT stock 5,168 / in transit 432,270, KIBING
  2,931 / 430,189 (same shipments, by supplier instead of by sheet). User decision (Nov 2026): FLAT stock stays with Subanci, also once
  the FLAT HK glass for Encap is in stock (MB52 only says FLAT) - do not share it with Encap.
  Downstream (step 2 on the aligned file): Inbound `grp_mats` / `name2grps` (deliveries of both sheets = 'G12R Back Glass'; totals
  conserved 1,529,059 = 1,049,389 + 479,670), `pool_g` (Summary / By Week: balance = total, days = total / total demand, periods per
  sheet), 'PPC In Transit' Shown on = the sheet with the PO / BL#; MRP_Daily Check `_sheet_pools` (stock vs balance, plan vs demand
  compared over the sheets together). Nov 2026 Check: the merge artefacts are gone and BL PGZS01588110 now shows WSA 15-Nov vs MIS
  24-Nov (was hidden on a merged row).
  **Sheets 'Stock' and 'Total_Plan_Req' are left out** of the aligned file (user, Nov 2026: not needed; `WSA_DROP_SHEETS`): removed from
  the workbook, references in hidden old sheets -> #REF! (Excel's own behaviour); the WSA file is now recognised by Total_Plan_Req OR
  PPC_Plan_Req (`_WSA_MARK`); plan end / material names / codes / UOM come from the MIS 'Total Plan Req' when the WSA file has no
  Total_Plan_Req (same content - checked Nov 2026). Two-block sheets are split (`WSA_SPLIT_NAMES`: 'Potting Material B', 'Tapes Perforated';
  labels kept via `LABELS`). Sheets open at today's row (frozen at D8), autofilter on row 7, red for Balance / days < 0.
  - Rows before today: values copied from the input (Balance style yellow if the input cell was yellow = typed).
    Today and later: formulas. Balance today = typed own stock (MB52 / PPC count rolled forward, JW out) -> the script's stock date
    = today; then C(r) = C(r-1) - Actual(r-1) + Delivered(r-1). D / E / customers / MW = SUMIFS on `PPC_Plan_Req` (block id = label, column
    header, date; 'Counted in Balance' = No for OEM). Plant Days = C/E. Plant + Port Days = (C + Scheduled row 2 + SUMPRODUCT(Scheduled 8:r, Remarks
    without 'canc', Delivered blank)) / E - the planner's own logic: Scheduled Delivery holds OPEN deliveries only (the quantity
    is moved to Delivered on receipt - no row has both), late open ones still count; cancelled ones no longer count.
  - Scheduled Delivery row 2 (X2) = SUMIFS(Archive Scheduled, sheet, State 'Not delivered'): open deliveries of deleted rows (8 in Oct-26).
  - Delivery columns copied by header (EPE 420's header-less Booking column handled); the planner's constant formulas (=10000*4) and
    same-row formulas (=A747-40) are moved to the new columns (`_carry_formula`); others become values. Header-less helper columns
    (old status / BL split formulas) are dropped (Status is rebuilt in AD).
  - Archive sheet (accumulates, deduped): delivery records of deleted rows, State Delivered / Cancelled / Not delivered.
  - Dashboard: column A index -> 19 for every redesigned row, sheet ranges -> $A:$AG, Potting B / Perforated rows -> the new
    sheets; a row reading two sheets uses the sheet named in column B (fixed: 'EPE 420 GSM' D18 read 'EVA 420 GSM').
    Hidden old sheets (Summary, Summary Dashboard, flux, EVA 480 GSM ...) that point at planning cells by address are not updated.
  - XML level: replaced sheets lose their old rels (comments / drawings / printer settings were all on deleted rows); local
    defined names of redesigned sheets removed, others renumbered after inserting the split sheets; new styles appended once
    (`_Styles`, deduped - styles.xml identical on re-runs); calcChain removed, fullCalcOnLoad set; every formula has its cached
    value (SUMIFS / SUMPRODUCT / status mirrored in Python) so the script reads the numbers without Excel.
  - Proof used (Nov 2026): every sheet has identical header / info rows; LibreOffice's own recalculation = cached values on all
    292,246 cells (planted wrong caches were corrected, so it really recalculates); simulated next day (rows before today identical)
    and month change (November rows -> Archive) pass; `--no-align` output = old script.
- **PPC In Transit scenario** (Nov 2026 user request: "another Plant + Port Days under my in transit"): `read_ppc_in_transit()`
  reads the PPC sheet `In_Transit` (MIS In Transit columns + the planner's own; stops after 500 empty rows - the sheet is formatted
  to row 1,048,000): qty = Final Pending Qty. (else Final Qty - PPC GRN), date = Final ATA2 / Final AtA (= Revised ETA Port + 5).
  MATERIAL = traced to the user's BOM (user rule, Nov 2026: "trace supplier name and description to my BOM code and description, and RM
  Code as well, because Final RM Code is sometimes wrong"): `in_transit_catalog()` = PPC BOM (code, description, supplier) + MIS
  Common Name / MB52 / MRP descriptions; `map_in_transit()` scores candidates (given codes, same code family, codes of the same
  supplier) by `_desc_score` + a -3 penalty per contradicting size / thickness / HS / GSM + product type (`_kind_of` = FIRST product
  word, so 'EPE FILM - EPE (EVA ...)' is EPE) + supplier match (`_sup_tokens`: no place names / 'CO LTD'; `SUPPLIER_ALIAS` KB =
  KIBING) + code evidence (RM Code 1.5 > Final 1.0 > Match / PO-wise 0.5). Final RM Code replaced only when another code is 2+
  points better; near-ties (< 1) go to RM Code, then Final; a line with no code needs a strong specific match, else 'not a BOM
  material'. YOUR WSA FIRST (user rule, Nov 2026: "you find PO, HBLNo., MBL in In Transit and PO + BL# in WSA - map this way"):
  `extract_fast` keeps PO + BL# of every dated row of every block ('refs') + the Archive sheet of an aligned file (by block label);
  `wsa_ref_index()` / `wsa_find()` = rows with the PO AND an HBLNo./MBL number (suffix match: MBL 'CMDUCHN4068035' = 'CHN4068035') >
  the BL number > the PO; the codes of the matched sheets that fit the description (same `_kind_of`, no contradicting size / th /
  HS / GSM / front-back) decide (several: best score); else the scoring above. CHECK (red, console line) only when the WSA sheet has
  the SAME part (type + side) in another size / HS = booked as another variant: Nov 2026 = TNBS glass 2272 (description + RM Code)
  booked on Glass G12R Topcon SB (2376, PO 6000005131), YONZ 35HS (G-G) frame booked on Frame G12R (30HS). Back glass 333 of PO
  6000005271 (Subanci) is on the planner's Subanci sheet = customer split, NOT a check: its arrivals go to 'Glass G12R Topcon
  Subanci Back' (`_wsa_shared`). Nov 2026 result: 246 lines matched on PO + BL#, 30 on BL#, 4 none;
  20 traced to another code than Final (17 earlier + 3 EP304 0.50 mm lines 024 -> 045: the planner's sheets have EP304 0.50 mm on
  EPE 420 GSM and 0.60 mm on EPE 480 GSM; EP304 LGCLAX69013X with no thickness = 420 GSM by its PO + BL#), 7 CHECK, 4 not BOM
  materials. Inbound sheet 'PPC In Transit' = the trace of every line (amber = traced to another code; red = CHECK / not BOM;
  column R = your WSA sheet with this PO / BL#). MRP_Daily uses the SAME tracing (user question Nov 2026: "did you do the same
  with the MRP output"): `normalize_in_transit()` (every read of MIS 'In Transit', after `extract_fast` has set `WSA_REFS`) runs
  `map_in_transit` on the MIS lines (catalog + WSA PO + BL#; alias {}) and writes the traced code into 'Final RM Code'; MIS In
  Transit = the same 280 shipments / 7,047,122 pending as the PPC In_Transit, and every mapped line now has the same material in
  MRP_Daily as in the WSA / Inbound. `IT_FIX` -> Check 'MIS Final RM Code does not match the shipment - traced code used' (detail
  ends with 'how: ...'): 21 lines (18 earlier + the 3 EP304 0.50 mm lines -> EPE 420; EPE 480 'Difference (MIS - WSA)' 400,239 -> 1.8).
  `ppc_arrivals()` = per MRP code and day, alias / factor like MIS
  (RMTTT00000018 = 2 x RMTTT00000010), late lines on today, 'No Material Found' lines skipped (console). Stored in
  `align_capture()['ppc']`. Reading is robust (Nov 2026: the user saw column T empty on EVERY sheet - not reproduced with the uploaded
  files, so the cause is on their PC): sheet found by name variants (`_it_sheet`: 'In Transit', 'IN-TRANSIT'), header row = first
  row with 2+ known headers (title rows allowed), numbers / dates typed as text read (`_num_cell` / `_date_cell`). When nothing
  arrives, `ppc['why']` says why (no PPC file; no In_Transit sheet + which other PPC file has one; quantity formulas with no saved
  result -> open + save in Excel; ...): T3 of every WSA sheet ('PPC In Transit: EMPTY - ...'), Inbound sheet 'PPC In Transit', and a
  WARNING as the LAST console line. Normal last line: "Column T 'PPC In Transit (arrivals)': N lines from <file> (sheet ...) -
  arrivals on k of n planning sheets" (`ALIGN_STAT`). `_ppc_arrivals_safe`: an In_Transit error no longer stops the alignment. WSA columns T 'PPC In Transit (arrivals)' (values),
  U 'PPC In Transit (Supplier Name)' (user request Nov 2026; suppliers arriving that day, joined ' / '; `ppc['sup']`) and
  V 'Plant + Port Days (PPC In Transit)' = IFERROR((C + SUM(T today..r)) / E, "na"); Delivered = W. `_hdr_role` 'ppcsup' is checked
  before 'ppcit' (both start 'ppc in transit'), extract_fast 'ppcit' skips 'supplier'. extract_fast reads 'ppcport' / 'ppcit' ('plantport' must not
  match '(ppc'); `below_zero_periods` key 'pi'; Inbound Summary T/U + Below zero rows (next delivery = next PPC arrival).
- **Other BOM materials** (no WSA sheet, Nov 2026 user request "don't create new sheets, show in Inbound"): `CAP[0]` =
  align_capture of EVERY run (also --no-align); after MRP_Daily, `other_materials_calc()` runs the WSA rules for MRP materials not
  on a WSA sheet (6 RM + 14 packaging today): Balance = stock - demand of the days before, Plant Days, Plant + Port Days (PPC In
  Transit); `add_other_materials()` reopens the Inbound file (openpyxl keeps the By Week chart) -> sheet 'Other materials' + rows
  in 'Below zero' (Planning sheet '(no WSA sheet)') + Summary D4 note. Solar cells are not in the MRP -> not listed (user told).
  Pine wood pallets show stock 0 from the 05-Nov count of 0 (a real count - user confirmed).
- **Inbound 'Below zero'** (Nov 2026 user request): `below_zero_periods()` = from the sheet stock date, periods where the sheet's
  Plant Days < 0 and where Plant + Port Days < 0 ('na' days keep the previous state). Summary columns R/S ('... < 0 (from -> to)',
  `periods_text`) + sheet 'Below zero' (one row per period, with the next delivery on/after From).
- **Units stay as SAP / MIS / PPC BOM give them** (user decision Nov 2026: "keep m2" for EVA / EPE). A WSA sheet label of another unit
  group ('Pcs' on EVA 420 / EPE 420 GSM; pc / roll counts as the same) shows the MRP unit in C4 (`ufix`, the numbers are in that unit),
  and a 'Supplier & Material' text of another product type (`_kind_of`: 'Frame 72 cells' on those two sheets) shows the material
  description (`mfix`, `align_capture` 'desc'); console 'Labels fixed ...'; the aligned file keeps them, so reruns are identical;
  MRP_Daily Check lost its 2 'WSA sheet unit vs SAP unit' warnings. Cell Fixation tape: PPC BOM has RMTTT00000018 (100 m)
  0.015-0.018/module -> RMTTT00000010 (50 m) x2 = 0.03-0.036; the old sheet used 0.008 - user answer Nov 2026: PPC is right (no change).

- **MIS dates are always the most updated** (user rule, Nov 2026, answer on BL PGZS01588110: WSA 15-Nov vs MIS 24-Nov). Every computed
  scenario already uses MIS dates: MRP_Daily (MIS In Transit), WSA column T / V and the ALERT (PPC In_Transit - checked Nov 2026: all
  287 lines have the MIS Final AtA). Only the planner's typed WSA delivery rows can be older (column S Plant + Port Days and the
  Inbound list follow them - the planner's own schedule). So both outputs tell the planner to move those rows: MRP_Daily Check 'Same
  shipment, different arrival date' -> "MIS is the latest (your rule): move the WSA row to dd-mmm" (15 BL groups on 07-Nov); Inbound
  column AQ 'MIS arrival (Final AtA) - the latest' when 3+ days from the plant date (`_mis_bl_dates` / `_mis_date_of`: the row's BL#
  = HBLNo. / MBL, suffix match, earliest MIS date), Chase? formula + To chase reason "MIS arrival dd-mmm (+N d) - MIS is the latest:
  move the WSA row" (35 deliveries on 07-Nov, 25 of them new on To chase), console 'To chase: ... MIS arrival 3+ days ...'. The
  planner's rows are NOT moved by the script (offer that only if the user asks).

- **Settings ALLOWED VENDORS in every projection** (user, Nov 2026: "Wsa output must follow settings, for example I don't see it follow
  allowed vendors by customer in settings. Also daily material status all sheet must follow settings"): `vendor_pools(dem, stock, arr)`
  = ONE rule for the WSA sheets (align_wsa, per sheet: its customers, stock by supplier, PPC arrivals by supplier; 333 via `_wsa_shared`
  'arr_v') and MRP_Daily (`_mrp_vendor_pools`: b['cust_line'], b['sup'], MIS shipments; result in VP_MRP, S minus lock). Unknown
  vendor = 'always allowed' (Settings text). Console: 'Settings ALLOWED VENDORS on 6 planning sheets ...' + '... projection by allowed
  vendor ...'. A material on 2+ WSA sheets (333): MRP_Daily takes the split from the aligned sheets (`_wsa_sheet_series` /
  `_vp_from_sheets`: short when any sheet is short) - checker finding: MIS shipments alone gave Encap's FLAT HK glass (PO 6000005517) to
  Subanci. Planned orders use S_pool - min(gap, lock) (= S when the pooled stock is already negative; checker round 2); supplier rows (Action List) use their own sheet / pooled S.
  Nov 2026 data: no WSA day changes (KIBING alone covers SB Energy, FLAT HK covers Encap); MRP_Daily 333 pooled row now = the ALERT
  Subanci Back periods (12-Nov, 24-28 Nov, 08-Jan on) - was 17-Jan: RM Issues Medium -> High; KIBING supplier row 22-Jan (= SB Back).
  Planted test (SB Energy 312 -> FLAT): SB Front V < 1 from 10-Nov, ALERT 312 red in all 3 workbooks, LibreOffice 'Check vs ALERT' 44/44 OK.
  Not done: a customer with 2+ allowed vendor rows uses its FIRST vendor only (as MRP_Daily supplier rows; none today).
  Not done (ask first): Settings SUBSTITUTES on the WSA sheets (EPE 480 / EPE 420 stay separate sheets; ALERT pools them).

- **'In transit supplier name (qty)' on every Daily_Material_Status sheet** (user, Nov 2026: "Add in transit supplier name in all sheet
  of daily material status"; `it_sup_text` / `_add_it_sup_col`, end of `make_daily_material_status`, DMS only - MRP_Daily unchanged):
  right after 'In transit' (Materials Status, ALERT, BOM Weekly), before the day columns (BOM Daily, RM dashboards; freeze F5 -> G5).
  'SUPPLIER qty; ...' largest first, legal form cut (`_sup_short`); code rows = MIS In Transit of the code (substitute group row =
  members); 333 per WSA sheet = PPC In_Transit lines of the suppliers in that sheet's column U (`ppcsup_names`). Check Nov 2026: the
  quantities add up to 'In transit' on every row of the 3 sheets that have it; all other cells unchanged.

- **RM Material Dashboard - MIS finds column V by its header** (user screenshot, Nov 2026: day cells showed supplier names, 0.0, 6E+09):
  `_wsa_dashboard_to(use_v=True)` read the fixed column 22; with 9+ customer columns (WSA_SLOTS is a minimum) V moves right. Now `vcol`
  = header 'Plant + Port Days (PPC In Transit)' of each sheet. Proof: WSA_SLOTS forced to 9 -> old 2,816 wrong cells, new 0; all three
  workbooks cell-identical to the 8-column run.

- **Dashboards in a light theme** (user, Nov 2026: "Change all Dashboard background from black to better color"): BOM Material
  Dashboard - Weekly / - Daily (both outputs) and the 2 RM Material Dashboards: page DASH_BG F4F6F9 (was 12131C black), text DASH_TXT
  1F2D3D, muted DASH_MUT 5D6D7E, tab DASH_TAB 2C3E50, 'na' DAY_NA DDE2E8 (was 1B2335), Short from in C0392B; the cover colours
  (DAY_BINS) and the navy group bands unchanged. Values cell-identical; no black cell / white text on the light page left.
  The supplier column takes the plain item-name style on the daily / RM dashboards only (plain=True; it copied the heat cell next to
  it); legend 'na' chip text 5D6D7E (white was invisible on DDE2E8 - checker finding).

- **Dashboards BLACK again** (user, Nov 2026: "output is messy, change background back to black as before"): DASH_BG 12131C, DASH_TXT
  E5E8EE, DASH_MUT 8A949E, DASH_TAB 12131C, DAY_NA 1B2335, Short from FF8A80 (`DASH_RED`), legend 'na' text white - the values before
  the light theme (which is the rejected one: do not offer it again). Values cell-identical. The supplier column's empty cells (title,
  band, No Plan, legend rows) take the item-name cell's fill (`_add_it_sup_col` plain / dark; checker finding: white holes on black).
  The 'messy' screenshot (PO numbers 6E+09 / supplier names / 0.0 in the day cells of RM Material Dashboard - MIS) = a run whose WSA
  file was NOT rebuilt (`--no-align`, or the alignment failed, e.g. file open in Excel): no column V header -> the old fallback read
  fixed column 22 of the planner's old layout. Now no fallback: a sheet without the header stays blank, and when NO sheet has it the
  sheet is not built (console NOTE 'RM Material Dashboard - MIS not built ... close the WSA file in Excel and run again').

- **Daily_Material_Status dashboards: no freeze, no wrap, one grid** (user, Nov 2026: "Dont freeze or wrap text, i want all dashboards
  in daily material status look organized" - read as: no cell under the dates with another width / height): `_tidy_dashboard` (end of
  `make_daily_material_status`, the 4 dashboards, DMS only - MRP_Daily's BOM dashboards unchanged): freeze panes off, wrap_text off on
  every cell, every row except header row 4 = 15 high, every day / week column one width, text columns as wide as their longest text
  on one line; the supplier column is capped at 48 (113 pushed the day columns off the screen - checker finding; 2 rows with 4
  suppliers are cut on screen, full text in the cell). Values cell-identical. Do not add frozen panes / wrapped text to these sheets again.
  Removing a freeze = also reset `sheet_view.pane` + `selection` (openpyxl leaves `<selection pane=..>` without a pane -> Excel
  'Repaired Records: View from sheet4.xml / sheet5.xml', user report Nov 2026). `_clean_views` before the Daily_Material_Status save =
  one selection per pane on every sheet (the hidden Materials Status / ALERT had 5; checker finding).
- **WSA file open in Excel = the 'messy' run** (user, Nov 2026: "this messy happened after you tell me to choose option 2"): since ONE
  WSA file (option 2) the rebuilt file is saved OVER the planner's file; open in Excel -> Windows blocks the save -> step 2 runs on the
  old layout. Before option 2 the separate '(aligned <date>)' file was written anyway. Now the run ends with a boxed 'WARNING: YOUR WSA
  FILE WAS NOT UPDATED IN THIS RUN ... CLOSE the WSA file in Excel and run the script again' (tested with a simulated PermissionError).
  Offered, not built (two WSA files on those days): save under a second name when the file is locked.

- **BOM Weekly dashboard: Stock, SAP stock (MB52), Physical count** (user, Nov 2026: "BOM Material Dashboard, add these: STORE · stock"):
  first 3 entries of `HEAT_MS_COLS` (E..G, before 'In transit'; weeks now from N), copied from the Material Status row by `_heat_ms_fill`
  like the other Material Status columns; both outputs. Check Nov 2026: 45 / 45 rows = Material Status, old columns unchanged.

## Known data quirks (NOT script bugs)
- WSA sheet `Glass G12R Topcon Subanci` (planner's original) lists only RMGLS00000356 but its Demand / deliveries count front + back
  glass. MIS splits it into 356 (front) + 333 (back). Since Nov 2026 the aligned file splits it too ('Glass G12R Topcon Subanci
  Front' 356 + 'Glass G12R Topcon Subanci Back' 333); on the planner's original file the Check sheet "Plan" mismatch for glass is expected.
- Nov 2026: the user quoted SAP closing stock 68,613 (268) / 72,417 (280); the uploaded MB52.xlsx and the PPC BOM_Usage sheet both say
  21,717 / 17,913 - their numbers come from newer files. 268 = back glass (printed, with hole), 280 = front (HTAR) per MIS Common Name.
- Customer spelling `Clenera Crimson` vs `Cleanera Crimson`: fix it on the Settings sheet.
- "In MIS, not in WSA" shipments on the Check sheet = the planner has not added them to the WSA sheets yet.

## How to work on this repo
1. **Get a concrete symptom first**: material code, sheet, value shown, value expected. Don't guess at fixes.
2. **Reproduce with the real input files** (the user uploads them; they are not in the repo).
3. **Prove the fix**: run the script before and after, then compare every sheet of the outputs. Only the expected cells may change.
   The console output should stay the same unless the fix is meant to change it.
4. Keep fixes small and explain them in plain language, with a before/after table for the affected materials.
5. **Never commit the Excel input or output files** (company data).
6. **Independent check before every reply that sends files** (user rule, Nov 2026: "I want an AI bot that checks your work; you just
   fix what it finds"): after testing a change, run the `output-checker` subagent (`.claude/agents/output-checker.md`; tell it the
   before-script, e.g. `HEAD~1`). Fix every proven finding, re-run the checker, at most 3 rounds. Send the files only on PASS, or say
   plainly which findings are still open. Show the checker's verdict and findings table in the reply.

## Working with the user
- Correct the grammar and vocabulary in the user's message first (US English), then restate the request as a developer requirement, then answer.
- **Two daily outputs, one logic** (see the top of this file): a rule asked for the WSA daily goes into MRP_Daily in the same version
  (and the other way round); prove both outputs in the comparison.
- **After every change to `update_inbound.py`, send the updated file to the user (attachment) in the same reply**, so they can run it.
  Send it only after the change is tested, and say in one line what changed in that version.
- **Also send the output files** (`MRP_Daily_<date>.xlsx` and `Inbound_<date>.xlsx`) from a fresh run of that script on the
  user's latest uploaded inputs. Put them in `output/` (git-ignored - company data, never commit or push them).
- Use simple, clear language. Planning terms (MRP, MB52, JW, OEM, BL, ETA) are fine.
- **Say when a new session is better** (standing user rule, Nov 2026: "tell me when you think I need a new session to continue with my
  project"). Say it at the end of a reply, with the reason, when: the conversation was already summarized once (context compaction) and
  the current task is done; a big new feature is next and the context is already long; or the next request is a new topic / a new
  day's run. Give a short handover: the work is on the session branch until merged into `master` (the default branch - a new session
  starts from master, so offer a pull request; create it only when the user asks), the 5 input files to upload (their WSA file - since
  Nov 2026 the script updates it in place, so it is the one file they keep working in), and the open points below.

## Open points (ask the user; remove when answered)
- None (Nov 2026: Cell Fixation tape = PPC BOM; Dashboard back-glass rows = yes, built).

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
To chase, By Week) and `<WSA name> (aligned <date>).xlsx` (redesigned planning sheets, see below).

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
  'ALERT (formulas)' keeps its own positions; its 'Check vs ALERT' points at the ALERT letters + 2 from H on. Group row = members
  x fx; 333 per WSA sheet = the suppliers whose estimated stock adds up to the sheet's stock (`_subset_match`, r['sups']; none = '-').
  Material Status 'Stock SAP' / 'Stock Count' -> 'SAP stock (MB52)' / 'Physical count' = the same values (Indirect Material sheet
  of Daily_Material_Status too). Stock Check: column B 'Material' -> 'Item name' (= ALERT, MIS Common Name) + C 'Category'
  (`item_category`), every later column +1; 'Estimated stock today' filled for every row (not counted = SAP); a material with stock
  from 2+ suppliers = one row PER SUPPLIER ('<name> - FLAT'): SAP = MB52 batches (`b['sup_sap']`, saved in `attach_wsa`), count =
  PPC count by Supplier Code (`b['sup_cnt']`, `apply_count_by_supplier`), estimate = the MRP Daily supplier row, usage / GR by count
  share, rest = 'Unassigned'; 333 supplier rows carry the ALERT label of their WSA sheet.
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
  Stock Management groups: A-F Sr. / Alert / Code / Item name / Category / UOM; STORE G-H Stock today, Days of stock; INVENTORY I Reorder
  level; QUALITY J-K Status, On hold; PRODUCTION L-M Need / day, Days incl. in transit; LEAD TIME N-Q; BELOW 21 DAYS R-T; WAREHOUSE U-X
  In transit, Next arrival, Overdue in transit, Arrival at plant. New fields from `alert_stock_fields` (in `build_alert_rows`, so the
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
  numbers; `align_wsa()` writes `<WSA name> (aligned <date>).xlsx` next to the WSA file (the original is never changed); step 2 =
  normal run on the aligned file (subprocess, `--no-align`). An alignment error is reported and step 2 runs on the user's file.
  The console ends with 'WSA ALIGNED FILE SAVED: <path>'.
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
  starts from master, so offer a pull request; create it only when the user asks), the 5 input files to upload (their OWN WSA file,
  not an aligned one), and the open points below.

## Open points (ask the user; remove when answered)
- None (Nov 2026: Cell Fixation tape = PPC BOM; Dashboard back-glass rows = yes, built).

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
| `MB52.xlsx` | SAP stock (Unrestricted + Quality Inspection). MIS MB52 gives names/keys only |
| `PPC*.xlsx` (optional) | **Module plan** = sheets `Plan` + `Revised Plan`; **qty per module** = `BOM`; physical count = `Count`; `In_Transit`; `PM_CS_Usage` |

Outputs: `MRP_Daily_<date>.xlsx`, `Inbound_<date>.xlsx`, `Daily_Material_Status_<date>.xlsx`, `WSA_Daily_<date>.xlsx`,
`Daily_Email_<date>.eml` / `.html`, `Run_Summary_<date>.txt`. Run: `python update_inbound.py` (Python 3.9+, pandas, openpyxl, lxml),
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
| In transit = same shipments, pending qty, plant arrival date (late lines on today) | PPC In_Transit -> column T | MIS In Transit (`ships`) |
| Material of a shipment = your WSA PO + BL# first, then description + supplier + RM Code | `ppc_arrivals` / `read_ppc_in_transit` | `normalize_in_transit` (both use `map_in_transit`) |
| Front / back glass = two materials | one sheet per glass (`_glass_split`) | per material |
| Supplier <-> customer = Settings ALLOWED VENDORS; stock by supplier = MB52 batches / count by supplier | 333 on 2 sheets (`_wsa_shared`) | supplier rows, `allocate_lines` |
| ALLOWED VENDORS on the projection (`vendor_pools`) | AK / AL columns, V subtracts the lock | `_mrp_vendor_pools` -> `x['S']` minus lock |
| ALERT list (`alert_eval`, `alert_why`, `build_alert_rows`, built ONCE) | Inbound first tab 'ALERT' (`add_alert_inbound`) | first tab 'ALERT' (`write_alert_sheet`) |

## Key rules (how the numbers should behave)

### Demand, stock, plan
- **Demand = module plan x qty per module**, by customer and line (not the WSA sheet demand). Plan = PPC `Plan` with `Revised Plan`
  on top (a Revised row replaces the Plan rows of its date + line; names fuzzy-matched; `ppc_plan_sheet()` / `read_revised_plan()`).
  Without a PPC file: `Production Plan` sheet of MRP.xlsx. Qty per module = PPC `BOM` for the customers it lists, else MIS Standard BOM.
  Plan customer -> MRP customer only through `LINKS` / Settings MIS LINK (`link_override()`), never exact spelling.
- **Stock = MB52**, replaced by the PPC `Count` sheet ONLY (newest date column with a number; a 0 there is a real 0), rolled forward
  (count - usage + GR since count). No PPC file = no counts (WARNING). MRP.xlsx / old MRP_Daily 'Stock Count' are never read for counts.
  PPC Count = its MAIN table only (`_count_main_table` cuts the second list at row 170).
- **BASE DATE = MB52 file date** when older than today (`base_d0()`): everything is calculated from it, a count dated on / after it is
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
- **Settings ALLOWED VENDORS in every projection** (`vendor_pools`): a customer tied to a vendor uses only that vendor + unknown-vendor
  stock; customers without a rule use everything. A customer with 2+ vendor rows uses its FIRST vendor only. 333 in MRP_Daily takes
  the split from the aligned sheets (`_wsa_sheet_series` / `_vp_from_sheets`). FLAT stock stays with Subanci (user decision).
- **Tapes: YOUR BOM CODE = THE MATERIAL** (`prepare_mis` `CODE_SWAP`): MIS tape code not in your PPC BOM while the BOM lists ONE other
  roll length of the same tape -> planned under the BOM code (RMTTT00000010 -> 018 x0.5, RMTAP00000025 -> 018 x2.5); MIS plan, In
  Transit and checks converted; the planner's typed WSA rows are already in the BOM code's unit (not converted). Item name stays the
  MIS Common Name (changing `b['common']` breaks the alias link).
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
  Dashboard colours `DAY_BINS(lt)`; Shortage Watch 'Short soon' = first short day within the lead time. Dead code is deleted (user:
  "delete all script that you ignore") - do not keep unused functions / constants.
- **Daily targets and OEM columns = Action List ONLY.** Every other sheet follows the module plan. Action List shows BOTH scenarios
  (do not ask again): daily TARGET ('Below 21 Days From' / 'Days Late' use `x['cw']`, with OEM target) and MODULE PLAN gold band
  (`x['cw_p']` = ALERT 'Below 21 days from' / 'Days late'; copied to MRP by Line). Status follows the plan. Stock-zero periods on the Action
  List = 'Short from -> to (with in transit)' only ('Short From (Stock = 0)' stays on RM Issues).
- **MRP Daily / Weekly**: 'Firm receipt (Open PO, no vessel)' and 'Days of cover (... + Open PO)' rows exist only for materials with an
  Open PO (rowmap `ro` / `rj` = None otherwise - guard every use); Supplier column = the single real supplier (`b['_sup1']`); supplier blocks
  only with 2+ suppliers (`multi_sup`). 'Inv. Days' rows = projected stock at day end / Usage/Day from that day.

### ALERT, Stock Management, Material Status
- **ALERT = first tab of MRP_Daily AND Inbound, same rows** (`build_alert_rows` -> `alert_eval`, stored in `ALERT_ROWS` / `ALERT_D0`).
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
- **'ALERT (formulas)' + 'ALERT Calc' + 'ALERT Plan lines' / 'ALERT Arrival lines'** (both outputs, `write_alert_formulas`, `alert_lines()`):
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
- **BOM Material Dashboard - Weekly** (`HEAT_NAME`, `add_cover_heatmap`, same rows as the ALERT) and **- Daily** (`DAILY_NAME`,
  `_bom_daily_sheet`, built LAST before `unify_workbook_style`): weekly value = the daily cover on that week's date (`bom_day_cover(r)[7w]`,
  colours `DAY_BINS`, short = below 1). Column A Code, B Item name, C Category; E..G Stock / SAP stock / Physical count and next columns
  from Material Status (`HEAT_MS_COLS`, `_heat_ms_fill`; EPE members = own values `own_ms`); 'Short from' in bold red. Final layout
  `_heat_final`: no title / note rows (header row 2, data row 3), 9 pt, columns E..'Days incl. in transit' outline level 1. Groups
  RM -> PM -> CS (`mgrp`, RMLAB checked first); a band row per non-empty group. Colours (MRP_Daily): softer dark (DASH_BG / DASH_TAB
  33383F, DAY_NA 474D57, bands 2B5A84, DASH_RED FF8A80). The all-light theme was rejected for MRP_Daily.
- **Daily_Material_Status** (`make_daily_material_status`): sheets 'Shortage Watch' (FIRST tab: Code, Item name, Category, Status Short now /
  Short soon (first short within `WATCH_SOON`) / Later / Covered / No demand, Days today, First short, Covered again, Short days; read from
  the BOM Daily cells, no new calculation, no line chart), 'Materials Status' (HIDDEN; Direct Material + OTHER MATERIALS band; Other Materials
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
  312 / 333; Subanci 356 / 333); quantities split by the BL#'s MIS In Transit lines (`_it_bl_index`), else 1:1; marker in A1 = never split
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
- **MIS dates are always the most updated** (user rule). Computed scenarios use MIS dates; the planner's typed WSA rows can be older and the
  script does NOT move them (offer only if asked): MRP_Daily Check 'Same shipment, different arrival date' ("MIS is the latest: move the WSA
  row to dd-mmm"); Inbound column AQ + 'To chase' reason when MIS is 3+ days off (`_mis_date_of`, suffix match).
- **Other BOM materials** (no WSA sheet): `CAP[0]` = `align_capture`; `other_materials_calc()` + `add_other_materials()` -> Inbound sheet
  'Other materials' + 'Below zero' rows. Solar cells are not in the MRP. Pine wood pallets stock 0 = a real count.
- **Inbound 'Below zero'** (`below_zero_periods()`): periods where Plant Days < 0 / Plant + Port Days < 0; Summary R/S + sheet. Glass 333
  deliveries of both sheets = 'G12R Back Glass' (`grp_mats`, `name2grps`, `pool_g`; MRP_Daily Check `_sheet_pools`).
- **ETD alert** (`ETD_LATE`, `write_etd_late_sheet`, sheet 'ETD alert' in BOTH Inbound and MRP_Daily): every open WSA delivery row (not
  cancelled; + Archive 'Not delivered') with an ETD whose BL# is not on MIS In Transit. `_etd_flag`: ARRIVES AFTER STOCK-OUT > ETD PASSED
  > SHIPS WITHIN `ETD_SOON` (7) DAYS > ON PLAN. A date in the BL# column = no BL# (`_bl_txt`). In no projection.

### Daily email (`write_daily_email`, `Daily_Email_<date>.eml` + `.html`) - the email the user sends
Dear all, intro (sources + base date), 'Needs attention' (2 bullets), lead-in line, 'Critical' table (Runs out in red, one period per line),
'ETD Alert' + reschedule sentence + table Material / Supplier / Deliveries / Quantity / Next ETD / WSA plant date, Notes, 'Best regards,
<name> / Production Planner'. Aptos 9 pt, no emoji. Tables split into Raw material / Packaging / Consumables (empty ones not printed; `is_pk`
= the `mgrp` rule). NO Summary section, NO counts in the subject / text, NO Overdue / Order-this-week tables, NO 'Alert' / 'Runs out'
columns on the ETD table; ETD rows = FUTURE ETDs that ARRIVE AFTER THE STOCK-OUT only, one line per material; facts only (no advice text).
Supplier (`_email_suppliers`) = in-transit suppliers (333 per sheet) > PPC BOM supplier (alias codes too) > MRP supplier-row stock
suppliers (glass / 2+ suppliers) > WSA Shipper > 'no supplier on file'. To / Cc / Name / Title = `Daily_Email_Settings.txt`
(`_email_settings`; company data - never in the script or the repo). Daily_Material_Status is attached; X-Unsent = Outlook opens a draft.

### Housekeeping
- **Old outputs are deleted** (`delete_old_outputs`, only after a full run saved today's files): earlier Daily_Material_Status, Run_Summary,
  ALL earlier Inbound files, MRP_Daily, emails, WSA_Daily. Exact script names only; `--keep-old` keeps all.
- `_WDT_CACHE` (WSA Dashboard read once) only inside the daily-file build and closed after (`_wdt_close`; open handles block the WSA_Daily
  replace on Windows).

## Known data quirks (NOT script bugs)
- MB52.xlsx has no save date inside -> the file date = the Windows file date (an uploaded copy gets today's date).
- WSA sheet `Glass G12R Topcon Subanci` (planner's original) lists only 356 but its Demand / deliveries count front + back glass (MIS: 356
  + 333); on the planner's original file the Check 'Plan' mismatch for glass is expected (the aligned file splits it).
- 268 = back glass (printed, with hole), 280 = front (HTAR) per MIS Common Name. SAP numbers the user quotes can come from newer files.
- Customer spelling `Clenera Crimson` vs `Cleanera Crimson`: fix it on the Settings sheet.
- "In MIS, not in WSA" shipments on the Check sheet = the planner has not added them to the WSA sheets yet.
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
   `/usr/bin/python3.11 -c "import ast;ast.parse(open('update_inbound.py').read())"`.
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
- SPEED (user, 12-Nov-2026: "The script is too slow"): 68-74 s on the real files (step 1 11-13 / aligned WSA 17-20 / step 2 39-43); no single
  hotspot (30 saves + 28 loads in step 2, `make_daily_material_status` ~1/3). Real gains need (a) dropping outputs / sheets the user does not
  open, or (b) a calc / presentation split. Asked which outputs they open daily + the 'Timing:' line from their PC (not answered).
- PPC Count second list at row 170 (24 packaging rows, stock units): newer count or scratch list? Ignored until answered; if newer it
  must be typed into the main table (pallets 535 / 447, 620, 603, 506 would change).
- Open (Low): RM Issues 'Project' blank for the EPE 480 + 420 group; Stock Management KEY RISK counts that group twice.

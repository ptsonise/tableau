# RULES_DETAIL.md - full rule detail for update_inbound.py

Read the section you need when a task touches that area. CLAUDE.md (repo root) holds the short version that loads every session.
This file = the full CLAUDE.md of 26-Nov-2026 (before the trim), unchanged below. New detail goes here; the one-line rule goes in CLAUDE.md.


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

Outputs: `1_MRP_<date>.xlsx`, `3_Inbound_<date>.xlsx`, `4_Status_<date>.xlsx`, `2_WSA_<date>.xlsx`,
`6_Email_<date>.eml` / `.html`, `7_Summary_<date>.txt`, `5_Web_<date>.html` (the ONLY web page). Run: `python update_inbound.py` (Python 3.9+, pandas, openpyxl, lxml),
~50-70 s. Every run also builds the 2_WSA file (step 1 = calculation + `align_wsa`, step 2 = normal run on that file, `--no-align`
skips, `--keep-old` keeps older outputs). Last console line: 'Timing: ...'.

**THE PLANNER'S WSA FILE IS NEVER CHANGED** (user 11-Nov-2026, option A). 2_WSA is a new file each day next to it; older
2_WSA files are deleted. `_wsa_tidy` removes script-made leftovers (exact names only; any other '(aligned ...)' file is kept with
a NOTE). If 2_WSA is locked in Excel: `_put_file` retries, then saves '... (2).xlsx' / '(aligned <date>)' with a boxed WARNING.

**TWO DAILY OUTPUTS - ONE LOGIC (standing rule).** WSA daily (aligned WSA file + Inbound) and 1_MRP must tell the same story.
Every calculation / mapping rule asked for one is applied to BOTH in the SAME change, without being asked; say in the reply how, and
show both in the before/after comparison. Only presentation choices the user placed on one sheet stay there. If a rule cannot apply to
the other output, say why and ask. Shared logic (one source, check both when changing):

| Logic | WSA daily | 1_MRP |
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
  (count - usage + GR since count). No PPC file = no counts (WARNING). MRP.xlsx / old 1_MRP 'Stock Count' are never read for counts.
  PPC Count = its MAIN table only (`_count_main_table` cuts a second list). Count sheet layout since 19-Nov-2026: the count column is
  'Physical Count Qty_BOM' (BOM unit; text 'no counting' = not counted) + ONE 'Count date' column (user 19-Nov-2026: "must read Physical
  Count Qty_BOM and Count date"). `_count_main_table` puts Qty_BOM / Quantity into the 'Count Closing Stock (BUN)' column renamed
  'Physical Count Qty' (MRP unit stays: RTV 006 plans in DR, 39,460 KG / 270 = 146 DR; CS / PM use Qty_BOM). 'SAP Closing Stock (BUN)' is not a count.
  Same Key + Qty_BOM + Count date as an earlier row = duplicate, counted once + NOTE (ribbon 036 rows 16/17). A material counted under
  2+ codes rolls forward from the OLDEST count date of the codes with stock > 0 (busbar 047 17-Nov; alias counted 0 on 19-Nov does not reset it).
- **PPC sheet `GR` = REFERENCE ONLY** (user 20-Nov-2026: "GR sheet is for reference only, don't count it in any calculation yet"):
  `read_ppc_gr` -> `GR_REF` (Qty in unit of entry on Posting Date; material via MIS MB52 Final Material Code + alias). Received since count
  stays 'PPC GRN' on 'GR. Date' of PPC In_Transit (`GR_LINES`; was MIS In Transit until 21-Nov-2026); in transit is NOT reduced by SAP GR (tried and reverted 20-Nov: SAP GR is a day
  ahead of MIS, so receipts were counted twice). 4_Status columns after 'Pending GRN Qty' (BOM Weekly in `DMS_WEEKLY_KEEP`,
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
  **Who is OEM = `oem_names()` / `is_oem_cust()` (src/11, ONE rule, user 28-Nov-2026 "8_Review mentions OEM plan customers - I had a rule to
  ignore OEM")**: Settings OEM = Yes + EVERY customer with modules on the PPC 'OEM Plan' sheet (`OEM_PLAN_CUST`, filled by `ppc_plan_sheet`;
  e.g. Banita, Desri, Rowdy Creek on Imperial Star / Arizona were in demand before = 38% of the plan) + the MIS customer each links to
  (LINKS, e.g. Desri -> Banita (Encap); name match also takes 'Banita (Encap)_IM'). Kept IN: a customer also on 'WSA Plan' (NOTE) and a MIS
  customer that a non-OEM plan customer links to. Used by the OEM step of `build_new_mrp` (-> ALERT Plan lines `OEM_IS`, 2_WSA) and `_plan_without_oem`.
- **Substitutes = SEPARATE materials on EVERY sheet** (user 25-Nov-2026: "Show material and its substitutions, e.g. EPE 480 and EPE 420,
  SEPARATELY on every sheet and dashboard: each with its own stock, plan, in transit and arrivals ... when calculating inv days show when and
  what quantity is covered by which material and its substitution code"). A Settings SUBSTITUTES code that is itself an MRP material (own
  block: EPE 480 RMEPE00000024 for EPE 420 RMEPE00000045, label 012 for 098) is NEVER folded into the other code (alias loop in `prepare_mis`:
  `if c_ in blocks: continue`) - not its stock, not its count, not its PPC In_Transit lines, not its WSA column T. Nothing is pooled any more
  (POOLG / group target / grp_* helpers deleted): Action List, MRP Daily / Weekly, MRP by Line, RM Issues, ALERT, dashboards = own numbers.
  `SUBG` (main -> [(sub, factor, use from, use until)]) feeds `sub_cover` (end of `build_alert_rows`): day by day, stock never below 0, the
  main code uses its own stock + arrivals, the substitute's stock + arrivals LEFT AFTER ITS OWN PLAN cover the rest -> `r['subcov']` (Days of
  cover per code, own cover to, substitute periods + qty, stock-zero periods with it) + `r['subnote']` (ALERT name note, email Situation).
  Shown: ALERT note, 4_Status BOM Weekly column `SUB_HEAD` 'Covered by substitute' after 'Days of cover (stock only)', web page Details 'Covered by
  substitute', email Notes. The substitute row sits right under its material on the ALERT (self-check exempt). Presentation only - classes,
  dates and projections are each code's own. 25-Nov: EPE 420 own zero 07-Dec (was 23-Dec pooled); EPE 480 counted 22,235 (was SAP 340,072 and
  its count added to 420), in transit 436,622 (2 x 218,311, 0.60 MM lines MSBV260229-3/-4, were on 420).
  Still folded (same material, other code): busbar 016 / 032 / 005 / 033 (not MRP blocks), tape roll lengths (CODE_SWAP), CS/PM '_Alt'.
- **CS / PM '_Alt' lines = the SAME material** (user 20-Nov-2026: "532 is alternate of 004, separately creates confusion"): PPC BOM 'Line #'
  = '<standard>_Alt' (PMPKM00000532 -> PMPKM00000004) -> `cspm_alt_fold` / `CSPM_ALT`: the alt line is used only when that customer / model
  has no standard line (never on top - it doubled the wrapping-roll demand), its count stock (`cspm_alt_stock`), MB52 and GR since count
  (`estimate_counted_stock`) go to the standard code; one row on every sheet, not offered on Substitute Advice. RM busbars 047 / 048 are
  two materials (both in every module); their Settings SUBSTITUTES (016 / 032, 005 / 033) are already in their stock (PACK_FACTORS alias).
- **Settings UNLIMITED MATERIALS empty = the script list `UNLIMITED_CODES` (523 / 338) stays** (user 22-Nov-2026); help sentences under
  UNLIMITED / ALLOWED VENDORS are not read as rows. MIS LINK overrides of customers not in this run's plan are still written to the Settings sheet.
- **Action List supplier rows** (glass per supplier + every material with a Settings ALLOWED VENDORS rule, user 28-Nov-2026: RTV 006 DOW / HU):
  'Usage per module by customer' = only the customers allowed that vendor (a customer with 2 vendor rows is on both); none ->
  'no customer may use this vendor' + Priority 'No demand' (312 FLAT). Settings name = batch Make by `_vend_hit` ('HU' = SHANGHAI HUITIAN, one row).
  Projection of a supplier row: 333 = its WSA sheet; dead supply (312) = pooled minus `dead_cum`; else = `vendor_pools` pools[vendor] (own
  balance; RTV DOW is used first by Pepper / Lucky 7, then HU), stock-zero floored from its own stock. Customers ('Lines' column) = `vendor_pools` who, else
  deduplicated (312 KIBING was 'SB Energy G12R, SB Energy G12R').
- **Supplier sub-rows on ALERT / Material Status / 4_Status / 5_Web** (user 28-Nov-2026: "a material with Settings ALLOWED VENDORS rules gets one
  sub-row per supplier under the material row (= Action List supplier rows). Numbers come from vendor_pools, no new calculation"): `vp_subs(r)`
  (src/07) = the Action List supplier rows of the ALERT row's material; `vp_sub_name` = '   ↳ <supplier> · <customers>'. No Code / Sr. on a
  sub-row, so check_align, the risk-order self-check, the ALERT usage column, 8_Review and the email skip it. Shown: stock (own, physical),
  in transit (own PPC lines), Stock zero from → to (`wins_txt` = the ALERT text rule; dead supply = 'not counted: no customer demand yet'),
  Arrival at plant (`ata_show` of its own arrivals with the material's Usage/Day = Action List). NOT shown: days of cover (no Usage/Day per supplier).
  ALERT: under the row; ALERT (formulas) compares with `ALERT_POS[i]` (the ALERT row of that material). Material Status: under the row
  (4_Status `_data_rows` skips sub-rows). 4_Status BOM Weekly: after the detail row, name / UOM / stock + one merged line 'in transit | stock zero |
  next:', same filter key as its material (`_dms_two_rows`). 5_Web: Summary tab, field `vs`. 333: none (its rows are already one per WSA sheet /
  supplier); 356 / 268 / 280: one supplier, none.
- **Settings ALLOWED VENDORS in every projection** (`vendor_pools`): a customer tied to a vendor uses only that vendor + unknown-vendor
  stock; customers without a rule use everything. A customer with 2+ vendor rows may use ALL of them, top row used first (user 23-Nov-2026 "Both allowed. Dow run out then use HU": `vendor_pools` = Hall's condition over groups of customers with the same vendor list; pools shown by order of use; RTV 006 Lucky 7 / Pepper DOW then HU, SB HU). Still first vendor only (no customer has 2 rows there today): `_wsa_shared` sheet split, glass supplier rows (`vend_of`), `bom_use_of` sheet filter. 333 in 1_MRP takes
  the split from the aligned sheets (`_wsa_sheet_series` / `_vp_from_sheets`). FLAT stock stays with Subanci (user decision).
  **DEAD supply** (user 18-Nov-2026 "why doesn't estimated stock follow?"): stock / arrivals of a vendor NO customer with demand may use
  (312 FLAT) are out of every CALCULATION from day 1 (`vendor_pools` `dead_st` / `dead_T`, in `lock` = WSA AK; `vp_dead(m)`): days of cover,
  stock-zero, Inv. Days, Action List days, Projected available balance, class, Stock Check 'Used in MRP' of that supplier row = 0.
  **SHOW IT, DON'T COUNT IT** (user 28-Nov-2026: "it is a physical fact ... ignore means don't count the stock and the arrival in calculation
  because it is no demand yet"; replaces the 18 / 22 / 28-Nov removal from the displays): `vp_nc(m)` = the not-counted stock / in transit + text.
  Shown physical: ALERT / Material Status / 4_Status / 5_Web 'Estimated stock today' (`stock_show(r)` = r['stock'] + nc; r['stock'] stays the
  counted number for every calculation), SAP stock / Physical count (`stock_src`, all suppliers), In transit (r['it']), Overdue, Action List Stock /
  In transit (its Inv. Days formulas subtract the not-counted part), MRP Daily / Weekly stock + In transit rows (the Projected available
  balance formula subtracts it per day, `NCd`), 'Arrival at plant' = counted arrivals + one line 'Not counted (no customer demand yet: <supplier>):
  in transit N (dd-Mmm → dd-Mmm)' (`txt_it`). ALERT 'Estimated stock today' has an Excel note with the full text. 2_WSA (user option A): column
  T / U list those arrivals; AK = the full lock (dead stock + arrivals), so V does not count them; the planner's C / R / S formulas do (pooled,
  not changed). 3_Inbound 'PPC In Transit': 'Counted on' = arrival date, 'Shown on' + '- not counted in cover / stock-zero: no customer demand yet'.
  Supplier row of a dead-supply material (Action List, e.g. 312 - KIBING) = pooled projection minus `dead_cum` (= ALERT). 3_Inbound 'Shown on' of a
  material on 2+ sheets = the sheet `_wsa_shared` puts the line on (`WSA_PUT`, passed to step 2 by env UPD_WSA_PUT). Check 'MIS plan vs WSA
  sheet demand': MIS side via `_plan_without_oem` (OEM out both sides). WSA Balance C / Plant Days R / Plant + Port Days S stay pooled (planner's formulas; not asked).
- **Tapes: YOUR BOM CODE = THE MATERIAL** (`prepare_mis` `CODE_SWAP`): MIS tape code not in your PPC BOM while the BOM lists ONE other
  roll length of the same tape -> planned under the BOM code (RMTTT00000010 -> 018 x0.5, RMTAP00000025 -> 018 x2.5); MIS plan, In
  Transit and checks converted; the planner's typed WSA rows are already in the BOM code's unit (not converted). Item name = the MIS Common Name
  with YOUR BOM roll length (`swap_name`, set in `build_new_mrp` AFTER the WSA capture - earlier it breaks the alias / sheet link; user 20-Nov-2026:
  MIS names 010 and 018 both 'Cell Fixation Tap-T50' -> shown 'Cell Fixation Tap-100m').
- **Unit = your BOM unit** (section 'UNIT = YOUR BOM UNIT'); a WSA label with no known conversion ('Pcs' on EVA = M2, "keep m2") / product-type
  text of another group is corrected to the MRP unit / description (`ufix`, `mfix`).
- **Check 'Usage per module: MIS BOM vs your BOM'** (`bom_rate_checks`) compares per customer and uses an alternative code only when the
  main code is not listed. **Check 'Stock vs WSA balance'** compares own stock (no JW).

### Usage/Day and day thresholds
- **Usage/Day** (`x['pl_need']`) = average plan per production day (plan > 0) over the material's TRANSIT LEAD TIME (`usage_day(P, lt_of(b))`);
  OEM plan out, each substitute code on its own. 'Inv. Days' = stock / Usage/Day; 'Inv. Days incl. in transit' = (own stock + in transit) / Usage/Day
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
- **ALERT = no summary boxes (old rows 1-9), no class cards / definition note (old rows 65-70)** (user 27-Nov-2026: 'dont need info from row 1 to 9 and 65 to 70'): `ALERT_HR`=2, data from row 3; every reader uses `ALERT_HR`, ALERT (formulas) refs follow `AO`.
- **ALERT = first tab of 1_MRP** (NOT in Inbound since 21-Nov-2026; ALERT (formulas) / Calc / Plan lines / Arrival lines also 1_MRP only) (`build_alert_rows` -> `alert_eval`, stored in `ALERT_ROWS` / `ALERT_D0`).
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
- **Count by supplier = the SAME unit as the material's count** (user 26-Nov-2026 'Fix'; `pcb_` loop in src/03): Qty_BOM when its unit is the MRP unit, else the count's own unit (RTV: 'Today Stock Qty' in DR, not Qty_BOM in KG). Stock Check supplier rows add up to the total (no 'Unassigned' leftover); the supplier split (`apply_count_by_supplier`) follows, so drum sizes count (HU 270 kg, DOW 250 kg).
- **Stock Check**: 'Usage / day (next 7 days, roll-forward only)' is NOT the Usage/Day of the other sheets (it only rolls the count forward; header says so, user 26-Nov-2026). 'Shortage risk' = the ALERT label (`alert_label(code)`; RM Issues still uses `risk_info`); rows in ALERT order;
  columns Item name, Category, Estimated stock, Lines / Project Name / Usage per module by customer / Module (P..S). 'Received since
  count' = `gr_between(code, count date, today)` from PPC GRN on MIS 'GR. Date' (Excel number -> `_any_date`). The CURRENT MONTH band was
  removed (user) - do not add it back.
- **'ALERT (formulas)' + 'ALERT Calc' + 'ALERT Plan lines' / 'ALERT Arrival lines'** (1_MRP only, `write_alert_formulas`, `alert_lines()`):
  the user can recheck the ALERT with Excel formulas; every row is verified against `r['P']` / `r['Tp']` (a row that does not add up gets
  per-day projection lines + console NOTE); 'Check vs ALERT' compares each column with the ALERT's own letters; no cached values
  (fullCalcOnLoad); use the `_xlfn.MINIFS` prefix; the line sheets get no per-cell font (speed).
- **Stock Management ('Material Status' sheet)**: `add_stock_mgmt_dashboard`, rows = EVERY material in AL_ORDER (RM + PM + CS, also No
  demand), right after the ALERT; column F 'UOM' (`_add_uom_col`). The Action List stays in 1_MRP (built in memory, other sheets use it).
  All 1_MRP sheets are visible. Self-check skips Check / Stock Data / No Demand / Substitute Advice (not in shortage order).
- **Action List = RM Issues look** (title 20 pt navy, header row 4 light grey, Calibri 10; 'Priority' column = RM Issues rule: High = zero
  inside lead time, Medium = short later, Low = under 21 d, OK, No demand; no stock-source columns - they are on Stock Data).

### Dashboards
- **One risk order on every dashboard** (`bom_risk_key` on `bom_day_cover`; RM dashboards `_risk`): first day the daily cover (days the stock lasts from that day, `bom_day_cover`) goes below 1, then below 5, then lowest cover today; no demand last; inside the RM / PM / CS groups.
- **NO dashboard sheets in 1_MRP** (user 21-Nov-2026: "Daily mrp, delete all dashboard sheet"): both BOM dashboards are still built in 1_MRP
  (4_Status copies them), then deleted from 1_MRP right after `make_daily_material_status`.
- **MRP by Line** (`make_line_view_xlsx`): NO merged cells - the material columns (Component, Stock, module-plan Usage/Day ...) are
  repeated on every line row (user 21-Nov-2026 "Why Mrp by line empty": a later openpyxl save kept only the first row of a merge, so a line
  filter showed blanks). Dates in the Action List columns get 'dd-mmm-yy'.
  Column D 'Line stops because of' (user 21-Nov-2026) = 'code name' of the material(s) that stop the line on 'Line stops on'
  (`line_run_days` `stop_by`, scenario 't'); Component is column E.
- **BOM Material Dashboard - Weekly** (`HEAT_NAME`, `add_cover_heatmap`, same rows as the ALERT) (the **- Daily** sheet is deleted, see the last bullet): weekly value = the daily cover on that week's date (`bom_day_cover(r)[7w]`,
  colours `DAY_BINS`, short = below 1). Column A Code, B Item name, C Category; E..G Stock / SAP stock / Physical count and next columns
  from Material Status (`HEAT_MS_COLS`, `_heat_ms_fill`); 'Short from' in bold red. Final layout
  `_heat_final`: no title / note rows (header row 2, data row 3), 9 pt, columns E..'Days incl. in transit' outline level 1. Groups
  RM -> PM -> CS (`mgrp`, RMLAB checked first); a band row per non-empty group. Colours (1_MRP): softer dark (DASH_BG / DASH_TAB
  33383F, DAY_NA 474D57, bands 2B5A84, DASH_RED FF8A80). The all-light theme was rejected for 1_MRP.
- **DMS BOM Weekly columns = the user's own edited file** (user 23-Nov-2026 'format output same as it'; supersedes the 23-Nov screenshot layout): Item name, Component Description from BOM, Item Category, UOM, Estimated stock today, SAP stock (MB52), Physical count, Days of cover, Usage/Day, First stock-zero day, Item name (repeated, also on band rows), Stock zero from -> to, Arrival at plant, Pending GRN Qty, SAP GR since count, Last SAP GR, 16 weeks; zoom 90, no 'Source:' note. Look since 27-Nov-2026 = 5_Web page (see 'BOM WEEKLY LOOK' bullet at the end). Presentation only.
- **4_Status 'BOM Material Dashboard - Weekly' = TWO ROWS PER MATERIAL** (user 27-Nov-2026: "each header is its own column, too wide for A4"; changes the old 'one line per data row, 19 high'; this sheet only). `_dms_two_rows` (src/10) is the LAST step, after `_dms_info_capture` (rows shift).
  Row 1 (18 high) = numbers only: Item name, UOM, Estimated stock today, Days of cover (stock only), week chips. Row 2 (15 high) = detail line merged over A..last week column: `Stock zero from → to` cell text in red (or grey 'no stock-zero day') + `     |     ` + `next_arrival_text` in grey
  ('next: Sun 29-Nov: 16.7 days   |   Tue 01-Dec: 8.3 days | ... +5 more:  39.1 days  (14-Dec → 08-Jan)' = EXACTLY the 'Arrival at plant (Final AtA)' column text, `ata_show` (days of cover, next ATA_SHOW dates, '+N more' total, PO / PR ETA when nothing in transit; user 28-Nov-2026 "match the Arrival at plant column"; blank column = 'next: none'); `next_arrival_text` (src/09) = the SAME function the 5_Web Summary 'next:' line prints (`nxl` in the web data). No value is calculated in the sheet step; the matching ALERT row is found by code + stock-zero text + days (333 = 2 rows).
  The warning triangle U+26A0 in the detail line = its OWN rich-text run (font 'Segoe UI Symbol', amber F5A623, colors written as FFRRGGBB): Segoe UI has no such glyph and Windows Excel's fallback font drew it black while an iPhone shows its yellow emoji (user 28-Nov-2026). Any new text with U+26A0 on a sheet does the same.
  Every other column (Component Description, Item Category, SAP stock, Physical count, substitute, Usage/Day, First stock-zero day, repeated Item name, Stock zero, Arrival at plant, PO / PR ETA, Pending GRN, POD, STATUS, GR) keeps its values (check_align reads them) as ONE collapsed Excel outline group (+ button).
  Filter 'Days of cover < PRINT_MAX_COVER' = hidden last column 'Cover (filter key)' that carries the days of cover on BOTH rows: a re-applied filter shows / hides the material and its detail row together. Print = `print_rows_fit` unchanged (hidden columns count 0 wide): 90 % / 5 pages -> 60 % / 1 page on the 27-Nov run.
  `check_align.py` fields 'Weekly detail row: Stock zero from → to' (= ALERT) and 'Weekly detail row: next arrival' (= 5_Web) must PASS. Plus the independent check "4_Status 'next:' line = the Arrival at plant column text": word for word the real column cell of the same row. Do not add a third row per material or put calculations in the sheet step.
- **4_Status = TWO visible sheets** (user 18-Nov-2026, RM Issues added 19-Nov): 'BOM Material Dashboard - Weekly' + 'RM Issues' (MOVED here from 1_MRP on 19-Nov-2026: copied in `make_daily_material_status`, deleted from 1_MRP after it; your Owner / Status / Date Closed come from the newest 4_Status, `load_rmi_prev`) with columns `DMS_WEEKLY_KEEP` only
  (Item name, 'Component Description from BOM' (PPC BOM Description, `BOM_DESC`; user 19-Nov), Item Category, UOM, Estimated stock today, Inv. Days, Usage/Day, First stock-zero day, Stock zero from → to, 'Arrival at plant (Final AtA)' (replaced Next arrival, user 19-Nov)) +
  the week columns (`_bom_weekly_keep`); no other sheet in the file. Header row 2 = centered, wrapped, 42 high, autofilter
  (`_dms_weekly_header`, user 18-Nov-2026 screenshot) - the ONLY wrapped row; data rows stay one line. Shortage Watch / Materials Status / ALERT are built then deleted (15-Nov).
- **'ETA - PO' + 'ETA - PR'** (two columns, user 24-Nov-2026 'Segregate PR and PO ETA': PO next to AtA, PR after it; first request "Deep dive sheet pr_po in ppc then insert EtA next to AtA in bom material dashboard"; DMS BOM Weekly ONLY, columns right after 'Arrival at plant (Final AtA)', `PO_HEAD` / `PR_HEAD` in `DMS_WEEKLY_KEEP`; information - in no calculation, not in 1_MRP / 2_WSA): PPC sheet `PR_PO` (SAP purchase requisitions + PO schedule lines; header row 1; Delivery Date `yyyymmdd`) -> `read_pr_po` -> `PRPO_LINES`, text by `pr_po_text`. Per line (Deletion Indicator skipped): **PO part** = Quantity ordered minus SAP GR on that PO + material (PPC `GR` sheet, oldest delivery date first) minus PPC In_Transit pending on that PO + material (that one is under AtA); **PR part** = Quantity requested - Quantity ordered (no PO yet). Text per column (`pr_po_text(kind=)`): 'Date passed: total', next `ATA_SHOW` dates 'dd-Mmm qty', rest as one total. Material = MB52 Final Material Code + alias, qty x alias factor (like `read_ppc_gr`); a supplier no customer may use (`vp_dead`) is left out. Read inside `make_daily_material_status` (In_Transit lines must be loaded). ETA assumed = PR_PO Delivery Date; In_Transit 'Revised ETA Port' (shipments already on the way) was NOT used - offer it if wanted.
- **PPC of 25-Nov-2026 has separate `PO` + `PR` sheets (no `PR_PO`)** (`read_pr_po` reads either layout): PR sheet = Delivery Date + Purchase order (+ item); PO sheet = open qty 'Still to be delivered (qty)' (already net of SAP GR; its 'Delivery Date' column is EMPTY and its 'Deletion Indicator' is an age, not a flag). PO part = PR line with a PO, qty = min(ordered, PO sheet open qty of that PO item) minus In_Transit pending; open PO qty with no date = 'No date: qty' (PO column only).
- **'Arrival at plant (Final AtA)' = ONE text in EVERY output** (`ata_show(txt, codes, need)`, user 26-Nov-2026: "Arrival at plant (Final AtA) in 6_email should be same as 4_Status and 1_MRP"; the 4_Status rule of 25-Nov-2026 "leave blank if nothing on the way ... update info from eta - po" + "need PR ETA also" now holds everywhere): 1_MRP ALERT (-> 4_Status BOM Weekly, which copies the ALERT), Action List (-> MRP by Line), 5_Web card + Details, 6_Email. PPC In_Transit arrivals (`_ata_text`) with each quantity as days of cover (`ata_days`); none -> 'PO ETA: ...' from the DATED open PO lines, then '  ;  PR ETA: ...' from the DATED open PR lines (`pr_po_text`, `eta_days`); nothing on any = blank (no 'nothing on the way' text any more). `read_pr_po` runs in `build_new_mrp` right after `CAP[0]` (needs the traced PPC lines), before the first text. The ALERT box 'Nothing on the way' still counts rows with no PPC In_Transit arrival (`r['ata']`). Empty arrival cells stay blank on the BOM Weekly (no '-'). Information only.
- **'Pending GRN Qty'** (user 18-Nov-2026, DMS ONLY, information - in no calculation): PPC sheet 'Pending GR' (`read_pending_gr`: PO.No,
  Invoice No, Supplier Name) matched to PPC In_Transit lines on PO + Invoice + supplier first word (`pending_grn`); qty = the line's Final
  Pending Qty in MRP unit, material = the traced code (`map_in_transit`; one invoice can be EPE + EVA). Column after 'Arrival at plant (Final AtA)' (BOM
  Weekly, in `DMS_WEEKLY_KEEP`) / after the in transit supplier column (BOM Daily, RM dashboards). Unmatched rows -> console NOTE.
- **4_Status** (`make_daily_material_status`, history): sheets 'Shortage Watch' (FIRST tab, user 14-Nov-2026: the ALERT columns without Sr. / Alert + 'Item Category' after Category, same rows and order as the ALERT,
  read from the workbook's ALERT sheet, one-line text, no frozen panes / wrap; `_shortage_watch`), 'Materials Status' (HIDDEN; Direct Material + OTHER MATERIALS band; Other Materials
  = PPC Count materials not in the BOM, need = 'Material Req Qty per Day'), 'ALERT' (HIDDEN, no class cards), BOM Weekly (opens first),
  (BOM Daily / RM Material Dashboard / RM Material Dashboard - MIS are NOT built any more - see the 'NO HIDDEN SHEETS' bullet at the end). `_wsa_dashboard_to` works the WSA Dashboard VLOOKUPs out in Python (WSA cached values are stale) and is used ONLY by `sort_wsa_dashboard` (risk order `DASH_MAP`). Every dashboard has Code + Item name; 'In transit supplier name (qty)' column after 'In transit'
  (`_add_it_sup_col`, DMS only). Both Material Status / ALERT copies: `_remap_rows` (no group band, header row 1). Consumable Unit = PPC sheet `PM_CS_Usage`, or `CS_Usage` when the planner renamed it (Nov 2026; column 'Code' + 'Consumable Unit ...'). Usage-per-module column at
  the end of both Status sheets; Direct Material has no STORE / WAREHOUSE / PRODUCTION cards (1_MRP keeps them).
  **Look of DMS dashboards = the user's 'WSA Material Coverage' page** (`_coverage_light`, DAILY FILE ONLY): white page, THREE cover colours
  (red under 7 days `COV_RED_DAYS`, amber under 21 days, green covered, blank = no plan / na); then `_tidy_dashboard`: NO freeze panes, NO
  wrapped text, every row 15 high except header row 4, one width per day column, supplier column capped at 48. Do not add frozen panes /
  wrap back. (`_fix_views` in `_safe_save` drops selections of panes that do not exist - Excel 'Repaired Records'.)
- `sort_wsa_dashboard()` re-orders the 2_WSA file's own Dashboard rows at XML level in the RM dashboard order (`DASH_MAP`); formula
  refs follow the row; on a base-date run F1 = DATE(base) instead of TODAY().

### 2_WSA file (aligned planning sheets, `align_wsa`)
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
  matches; CHECK (red) only when the WSA sheet has the same part in another size / HS. 1_MRP uses the same tracing (`IT_FIX` -> Check
  'MIS Final RM Code does not match the shipment - traced code used'). Inbound sheet 'PPC In Transit' = the trace of every line.
  **Copied description = the unit price decides** (user 21-Nov-2026 'go'): 2+ lines of ONE invoice with the SAME description but
  DIFFERENT 'Unit Rate', each price with its own Final RM Code -> Final RM Code used, before the WSA / description step (`map_in_transit`
  step 0, `IT_COPIED`, console NOTE). 21-Nov: Kibing KB-WSA-260716B-5 / -6 (both lines 'HTAR', 4.02 = 312 front, 4.07 = 333 back) ->
  36,000 front + 36,000 back (was 72,000 front).
- **IN TRANSIT = PPC In_Transit ONLY** (user 21-Nov-2026: "Script find in transit from PPC only. Ignore mis in transit"): `_CachedWB`
  hands every reader of the MIS 'In Transit' sheet the PPC 'In_Transit' rows (`ppc_it_rows`, same columns as MIS): 1_MRP projection,
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
  asked): 1_MRP Check 'Same shipment, different arrival date' ("PPC In_Transit is the latest: move the WSA row to dd-mmm"); Inbound
  column AQ + 'To chase' reason when PPC is 3+ days off (`_mis_date_of`, suffix match).
- **Other BOM materials** (no WSA sheet): `CAP[0]` = `align_capture`; `other_materials_calc()` + `add_other_materials()` -> Inbound sheet
  'Other materials' + 'Below zero' rows. Solar cells are not in the MRP. Pine wood pallets stock 0 = a real count.
- **Inbound 'Below zero'** (`below_zero_periods()`): periods where Plant Days < 0 / Plant + Port Days < 0; Summary R/S + sheet. Glass 333
  deliveries of both sheets = 'G12R Back Glass' (`grp_mats`, `name2grps`, `pool_g`; 1_MRP Check `_sheet_pools`).
- **ETD alert** (`ETD_LATE`, `write_etd_late_sheet`, sheet 'ETD alert' in BOTH Inbound and 1_MRP): every open WSA delivery row (not
  cancelled; + Archive 'Not delivered') with an ETD whose BL# is not on PPC In_Transit. `_etd_flag`: ARRIVES AFTER STOCK-OUT > ETD PASSED
  > SHIPS WITHIN `ETD_SOON` (7) DAYS > ON PLAN. A date in the BL# column = no BL# (`_bl_txt`). In no projection.

### Daily email (`write_daily_email`, `6_Email_<date>.eml` + `.html`) - the email the user sends
Layout = the user's 18-Nov-2026 sample: Dear all, intro (sources + base date), 'Needs attention' (2 bullets, NO reschedule sentence - duplicate, user 18-Nov),
lead-in line, '1. Critical' + one line ('Materials will run out of stock within the transit lead time. ' + reschedule sentence), sub-tables
'1.1 Raw Materials' / '1.2 Packaging' / '1.3 Consumables' (columns Material / Supplier / Cover today / 'Stock zero from → to (⚠ = starts
inside the transit lead time)' = ALERT `r['per']` in red, no wrap / 'Arrival at plant (Final AtA)' = `ata_show` (= 4_Status / 1_MRP, user 26-Nov-2026) / Situation), '2. Upcoming ETD Alert' + one line
+ '2.x' tables Material / Supplier / Deliveries / Quantity / Next ETD / WSA plant date (numbers follow the sections / groups that have rows), Notes. NO signature (user 18-Nov-2026: Outlook adds it). Aptos 9 pt, no emoji. Tables split into Raw material / Packaging / Consumables (empty ones not printed; `is_pk`
= the `mgrp` rule). NO Summary section, NO counts in the subject / text, NO Overdue / Order-this-week tables, NO 'Alert' / 'Runs out'
columns on the ETD table; ETD rows = FUTURE ETDs that ARRIVE AFTER THE STOCK-OUT only, one line per material; facts only (no advice text).
Supplier (`_email_suppliers`) = in-transit suppliers (333 per sheet) > PPC BOM supplier (alias codes too) > MRP supplier-row stock
suppliers (glass / 2+ suppliers) > WSA Shipper > 'no supplier on file'. To / Cc = `Daily_Email_Settings.txt`
(`_email_settings`; company data - never in the script or the repo). 4_Status is attached; X-Unsent = Outlook opens a draft.

- **BOM usage on the web page** (user 22-Nov-2026): `bom_use_of(code, where)` (one source) = line 'BOM usage N UOM / module' on the row / card + 'BOM usage per module by customer (source)' in Details. Presentation only.

- **DETAILS like WSA on 5_Web** (user 22-Nov-2026 "have detail like wsa has"): Details also hold the projected-stock chart (`det_data`: ALERT `S` = stock + PPC arrivals - plan, and stock-only line, PPC arrival ticks, stock-zero day, `DET_DAYS`=112 days), table 'PPC In_Transit arrivals' (ALERT `Tp`) and 'Open WSA deliveries' (`WSA_DLV` from the 2_WSA sheets of the same material / `where`; empty when the material has no WSA sheet). Shared `DET_CSS` / `DET_JS` / `det_inject`. Presentation only.

- **Line allocation reads Settings vendor names like 2_WSA** (user 26-Nov-2026 'A yes'; `allocate_lines` `src_of` = `_vend_hit`): Customer Supplier / Vendor Priority 'HU' = supplier row 'SHANGHAI HUITIAN ...', 'KB' = KIBING. Before, an exact-name match left SB Energy lines 3/4 RTV with NO supplier and Pepper / Lucky 7 with DOW only: MRP Daily HU 749 vs 2_WSA 716 on 27-Nov. Now the supplier rows add up to the material every day. Two plan lines with the same line key ('Arizona' + 'Arizona -Desri') ADD their demand (was overwritten: RTV supplier rows did not add up from 01-Feb; line short dates 1-5 days late for RTV 006 / 019, busbar 048, frame 131, tape 025). Vendor Priority code removed (never filled: `vpri` was always empty).
- **Stock Check 'Received since count' per supplier = that supplier's own receipts** (user 26-Nov-2026 'B yes'; `b['sup_rec']` from `apply_count_by_supplier`): 333 FLAT 0 / KIBING 7,200 (was 3,078 / 4,122 by count share).
- **Action List supplier rows (glass with 2+ suppliers) = their OWN arrivals** (user 26-Nov-2026 'Own'): 'Arrival at plant (Final AtA)' of a supplier row = only that supplier's PPC In_Transit lines (`_vend_hit`), like its stock / in-transit columns; the material row keeps all arrivals.
- **CUSTOMER GOES WITH SUPPLIER = one rule at the SOURCE of every list** (user 22-Nov-2026, 4 complaints in a row: "just a rule customer go with supplier ... this error repeated for years"; fix it at the source, never patch one page): (1) 2_WSA columns T / U: since 28-Nov-2026 they SHOW arrivals of a vendor no customer of that sheet may use (312 FLAT on 'Glass G12R Topcon SB Front', user option A 'show it, don't count it'); column AK holds them (full lock), so column V does not count them; (2) web page: BOM usage = only the customers of THAT sheet (`bom_use_of(code, where)`, `_cust_vendor`, customers without a Settings daily target but listed in your PPC BOM are included - Pepper), 'Open WSA deliveries' / WSA arrivals tables drop shippers the card's customers may not use (`usage_vendors`, `vend_ok`). Any NEW list of arrivals / deliveries / suppliers must apply this rule before it is written.
- **Inv. Days incl. NEXT ARRIVAL ONLY** (user 21-Nov-2026: "Web page and pictures, I also want to see inv day with the next arrival only"):
  `alert_eval` `days_nx` = (Estimated stock today + qty of the first arrival in T, late lines on the base date) / Usage/Day. Shown on
  5_Web (Details, panel column 'Days of cover (+ next arrival only)').
  Presentation only: in no Excel output and no calculation.
### Housekeeping
- **Old outputs are deleted** (`delete_old_outputs`, only after a full run saved today's files): earlier 4_Status, 7_Summary,
  ALL earlier Inbound files, 1_MRP, emails, 2_WSA. Exact script names only; `--keep-old` keeps all.


### UNIT = YOUR BOM UNIT (user 26-Nov-2026 - BUILT 26-Nov-2026)
User: "I want to see unit of measure in my bom only. If you see other unit you must ask me how to convert into my bom unit, e.g. RTV 1 DR of HU = 270 kg, 1 DR of DOW = 250 kg".
Answers (26-Nov-2026): 1 RTV006 shown in KG (BOM); DR -> KG by supplier: HU 270, DOW 250 (= PPC Count 'Quantity' column per row) | 2 BOX -> PC = 1,000
(laser labels CST1710 / CST1334; SAP stock 7 / 65 boxes now 7,000 / 65,000 PC) | 3 EA = NOS | 4 NOS = PC | 5 LITER = L | 6 ROLL -> KG from Count
'Quantity' (14.45 / 21.8) OK as is | 7 not in your BOM (RMALF131, RMJUB167): keep SET (SAP unit) | 8 2_WSA follows too: the planner's drum-typed Sealant
sheet converted to KG (deliveries by shipper factor) | 9 unknown supplier of a DR line = 270 (HU, main) + flagged on Check | 10 a unit with no conversion =
run continues, line shown 'UNIT ?' on Check and LEFT OUT of the numbers (never guessed).
**Code (src/02)**: `bom_unit(code)` = PPC BOM 'UoM' (blank / 'None' = not in your BOM -> SAP unit stays); `sap_unit(code)` = raw MB52 unit (PPC, MIS), MIS
Common Name, PPC Count base unit; `to_bom(code, qty, unit, sup, src)` = THE conversion: same unit / synonym (`UNIT_SYN`: EA=NOS=PCS=PC, LITER=L, ROLL=ROL,
DRUM=DR) unchanged; else x PPC Count 'Quantity' of the row with that Base Unit + BOM UOM (`_unit_tables`, per 'Supplier Code'); supplier by `_sup_pick`
(one supplier = it; SAP batch '1HU0..' / name word 'HUITIAN' starting with the code; unknown = first Count row = main supplier, logged); none -> None.
Applied where quantities ENTER, once: `_CachedWB.__getitem__` -> `_unit_rows` (`UNIT_SHEETS`: MB52 [stock, batch = supplier], MIS Total Plan Req +
Standard BOM [rates, plan], Common Name [label], RM Status for OEM JW [JW], PPC PR / PO / GR [ETA columns, GR reference]; unit cells relabelled to your
BOM unit), `normalize_in_transit` (1_MRP: Final Pending / Final Qty / PPC GRN by the traced material + supplier), `ppc_arrivals` (2_WSA, same rule),
`wsa_block_to_bom` (src/13: the planner's WSA sheet cells Balance / Demand / Actual / Delivered / Scheduled / ... when the Count has the conversion of the
sheet unit, e.g. 'DRUM(270KG/DRUM)' -> KG; deliveries by Shipper; used by `_wsa_read` (2_WSA, Archive) AND `extract_fast` (1_MRP Open PO WSA); E4 says
'your sheet unit ... -> KG'). Count = Qty_BOM (its BOM UOM = MRP unit). MRP unit (`b['uom']`) = your BOM unit first. PACK_KG / BLOCK_INFO deleted.
**Check** (`unit_bom_checks`, src/05): one row per (source, code, unit, supplier, factor): 'Converted to your BOM unit' (Low; Medium when the supplier is
not known = main supplier used) and 'UNIT ? - no conversion, left out of the numbers' (High). Proof 26-Nov-2026: count 62.52 DR x 270 + 24 DR x 250 =
22,880 KG; stock rolled forward 18,036 KG (was 68.58 DR); in transit 1,820 DR = 491,400 KG; labels SAP 7 / 65 BOX = 7,000 / 65,000 PC; check_align PASS.

## 8_Review (4_Status challenge panel, user 28-Nov-2026)
User: "4_Status is the only report to the supply chain manager, production manager and S&OP manager to challenge my output report. 3 critical
most important questions only each person." `write_review_panel` (src/09), called after 5_Web: EXACTLY 3 questions per manager about what
4_Status shows; the planner's answer = ALERT rows (= BOM Weekly) + the finished 4_Status read back (`_review_status_read`: BOM Weekly stock
columns, RM Issues). Rule-based, offline. Separate file because 4_Status = exactly 2 sheets. Presentation only, no new number.
- **S&OP**: 1 commit the plan for `REVIEW_WEEKS` (4) weeks = class 0/1 with first stock-zero inside; 2 MW at risk + customers = open RM Issues by
  Production Impact (MW), Project (not added up: they overlap); 3 orders to release this week = class 1, first stock-zero - lead time <= `RELEASE_WINDOW`.
- **Production**: 1 line stop in `COV_RED_DAYS` days (class 0) + next arrival (`ata_show`); 2 substitute (`subcov`) for class 0;
  3 stock trust = SAP MB52 vs physical count more than `REVIEW_GAP` (50 %) apart, at-risk materials first.
- **Supply chain**: 1 class 0 with nothing in PPC In_Transit (+ PO / PR ETA from `ata_show`); 2 class 0 whose next arrival is after the first
  stock-zero day (days late); 3 overdue (`ovd`) + Pending GRN (`pending_grn_qty`).
- Scope (user 28-Nov-2026): S&OP + Supply chain = raw materials only (`_heat_grp`==0 = the 4_Status RM group; RM Issues are RM already); Production = all materials (RM, PM, CS).
- Every answer lists EVERY material and every customer (user 28-Nov-2026: no '+N more on 4_Status', no '...'; e.g. the Subanci glass 356 was hidden). REVIEW_* numbers = DEFAULT. Old 8_Review files deleted.

## Numbers register (added 21-Nov-2026; enforced by `check_rules.py`)
Every hard-coded threshold / fallback has an OWNER. **USER** = the planner decided it: never change it without asking first.
**DEFAULT** = nobody confirmed it: do not change it silently, do not quote it as the planner's rule, and ask the planner to confirm it
when it matters. A number NOT in this table = treat as DEFAULT and add it here (and to `check_rules.py`) when you touch it.
After agreeing a change: update the code, this table and `check_rules.py` in the SAME commit. `python check_rules.py` must PASS.

| Number | Value | Owner | Note |
|---|---|---|---|
| Every day threshold = the material's transit lead time | rule | USER (Nov 2026) | target days, ALERT 'under lead time', reorder, look-ahead |
| `COV_RED_DAYS` red cover | 7 days | USER 12-Nov | "Red is under 7 days" |
| Web pages: list rule (= the material's transit lead time, `in_lead_time`; was `DASH_HTML_MAX_DAYS` 21) / weeks (`DASH_HTML_WEEKS`, `WSA_HTML_WEEKS`) / colours / width | lead time / 16 (scroll sideways) / red <7 (incl. zero), light red 7-21, grey 21+ (all 3 pages, 22-Nov) / week column 56 px fixed, material column 2/3 | USER 22-Nov | both pages (WSA web page too) |
| `BOARD_MAX_DAYS` Shortage_Board day strip | 63 days | DEFAULT | strip = longest lead time of the listed rows, capped here |
| `CONSUMABLE_CODES` | 7 codes | USER | |
| `LEAD_TIME_RM` fallback | 49 | USER 21-Nov | only when the Settings row is missing (Settings row wins) |
| `LEAD_TIME_CSPM` fallback | 14 | USER 21-Nov | |
| `PLAN_AVG_DAYS` Action List daily plan | 49 days | USER 21-Nov | fixed in the code, NOT a Settings row |
| `CUST_TARGET_DEFAULT` | Pepper 3000, Lucky 7 3000, Crimson 3600, SB 3600, Snowflake 4000 OEM | USER 21-Nov | only when Settings has no CUSTOMER DAILY TARGET |
| `ALERT_HORIZON` / `RISK_DAYS` / `RELEASE_WINDOW` / `ETD_SOON` | 112 / 5 / 7 / 7 | USER 21-Nov | Settings rows |
| `RESCHED_OUT_X` / `SHORT_HORIZON` / `ATA_SHOW` / `MOVE_DAYS` | 3 / 60 / 4 / 14 | USER 21-Nov | |
| `SHORT_LEAD` / `LONG_LEAD` | 20 / 75 days | USER 21-Nov | |
| `DASH_STEP2` / `DASH_BLUE1` / `DASH_BLUE2` / `DASH_MIN_LT` | 14 / 14 / 39 / 14 | USER 21-Nov | 1_MRP dark colours |
| `PRINT_MAX_COVER` DMS BOM Weekly filter | Days of cover (stock only) < 22 | USER 25-Nov | rows hidden by the autofilter, print / PDF shows only these |
| `PRINT_MIN_SCALE` A4 print zoom floor | 55 % | DEFAULT 25-Nov | ALERT / Action List / BOM Weekly print (`print_a4`) |
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
- 4_Status 'Other Materials': PPC Count column K ('Material Req Qty per Day') is an XLOOKUP into the PPC 'Usage' table and has
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
   After the run: `python check_align.py` (must PASS) = every field shown in 2+ outputs (stock, SAP, count, Usage/Day, days of cover,
   lead time, first stock-zero day, stock zero from -> to, Arrival at plant, next arrival, in transit; 1_MRP ALERT / Material Status /
   Action List, 3_Inbound, 4_Status, 5_Web, 6_Email) has the same value for the same material. A DIFF = one output has its own copy of a
   rule: move it to the ONE shared function, never patch one output (user 26-Nov-2026).
6. **Independent check (user rule, Nov 2026, lightened 14-Nov-2026 to save usage)**: run the `output-checker` subagent
   (`.claude/agents/output-checker.md`; tell it the before-script, e.g. `HEAD~1`) ONLY for LOGIC changes (a calculation, mapping, projection,
   date, stock, plan or vendor rule). Do NOT run it for text, label, colour, layout or width changes - for those do step 3 only. Fix every
   proven finding and re-run the checker, at most 3 rounds. Send the files only on PASS, or say plainly which findings are still open.
   Show the checker's verdict and findings table in the reply. State in the reply whether the checker was run and why / why not.

## Working with the user
- Correct the grammar and vocabulary in the user's message first (US English), then restate the request as a developer requirement, then answer.
- **Two daily outputs, one logic** (top of this file): prove both outputs in the comparison.
- **After every change to `update_inbound.py`, send the updated file to the user (attachment) in the same reply**, after it is tested, with one
  line on what changed. Also send the output files (`1_MRP_<date>.xlsx`, `3_Inbound_<date>.xlsx`) from a fresh run on the latest inputs,
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

- **5_Web_<date>.html** (`write_material_daily_update`, named 'Material Daily Update' by the user 22-Nov-2026; was Match_Centre, user 22-Nov-2026 "new dashboard inspired by OneFootball match page" + 4 app screenshots): 4th web page, one material = one 'match' (list of cards -> match screen: dark scoreboard header Inv. Days today vs Inv. Days + 21 d arrivals, arrivals = goals, stock-outs = red cards; pills Overview / Arrival / Weeks (user 23-Nov-2026 "weeks view tab next to arrival": the same 16 week cards for this material only, `wkHtml(rs,one)`, always open) (BOM usage tab DELETED 23-Nov-2026 "move per module to overview details": Per module + customer lines are now a row in Overview Details; the "Arrival at plant" block of that tab is gone; user 22-Nov-2026: Stats + 'Will it last?' + bell deleted, Ticker renamed Arrival); Overview = 'Coverage' (was Recent form; user 22-Nov-2026; design = the Shortage_Board week strip, user 22-Nov-2026 'Apply this design': one box per week up to the longest lead time of the listed rows, green triangle = PPC arrival that week, day / month under it, weeks after the row's own lead-time week faded) = Inv. Days per week; no 'D+N' day offsets anywhere, dates only (user 22-Nov-2026 'D+ is confusing') (whole number rounded down, red < 7, light red 7-21, grey 21+), Details, Live table = risk rank with neighbours, Top arrivals; header 'vs' = the material name). Same rows / order as Material_Shortage; dark header only on the match screen (user's reference), page light. Presentation only; older ones deleted.

- **WEB PAGES = LIGHT THEME ONLY, WAAREE LOOK** (user 22-Nov-2026: "Keep light theme only. No more dark"; "Delete png, no need"): Material_Shortage / 2_WSA / Shortage_Board use Poppins, brand green (#2B7A3E, white page #FFFFFF), yellow #F5A623 accent, gradient bar + WAAREE wordmark; NO light/dark button, NO dark CSS. Chat_Update pictures / text were removed (`write_chat_update` deleted).

- **5_Web bars = ONE week strip** (user 22-Nov-2026 "fix this and I want all the bar look like this"): list cards, Coverage and Arrival tab all use `wkb(r)` (boxes + green triangle row + date); the old runway bar is gone. Triangle row is a flex box (line-height 0) so it never overlaps the date. Presentation only.

- **5_Web Arrival tab: 'Days of cover that day (after usage)'** (user 22-Nov-2026 "Bar tell red but below all green"; replaces "Days of cover (+ arrivals up to this one)", which did NOT subtract usage and showed 7.5 / 22.8 / 30.4 under a red week strip) = the daily cover on the arrival day, the SAME number the week-strip boxes use (`bom_day_cover` -> `r.cv[day]`, stock after usage incl. that arrival / that day's demand); below 1 = '0 · still short'; NOT capped at 60 (`ac`, output-checker finding); an arrival on a day without demand (Sunday) takes the next production day's cover + '(first production day dd-Mon)'. Presentation only. The scoreboard 'Days of cover (+ next arrival only)' and the 'incl. next arrival only' figures still ignore usage (user-defined, unchanged).

- **STOCK NEVER GOES BELOW 0 in every cover / stock-zero calculation** (user 23-Nov-2026 "How can you assume arrival can cover the zero inv in the past", then "go ahead with what I said"): demand lost on a stock-zero day is NOT owed; stock restarts from 0 at the next arrival. One source: `floor_path` / `floor_wins` (stock after each day floored at 0; stock-zero day = floored stock + net flow < 0). Used by `bom_day_cover` (week strips, daily cover, risk order, BOM dashboards, 4_Status), `alert_eval` windows (ALERT 'Stock zero from -> to', zero days in lead time, class), Action List / RM Issues `wins` ('Short from -> to'), supplier rows, `det_data` (Details chart), ALERT Calc formulas (Projected stock = MAX(0, ...), Stock zero = IF(prev + arrivals - plan < 0)), 2_WSA.html week cells of a sheet with an ALERT row, and the 5_Web Arrival tab ('Days of cover after this arrival (stock never below 0)'; 356: 29 Nov +30,000 = 4.7, 01 Dec +80,000 = 18.3). NOT floored (planner's own WSA formulas): 2_WSA.xlsx Balance C / Plant Days / Plant + Port Days / column V, MRP Daily / Weekly projection rows - ask before changing. Projection S itself (`x['S']`, 'Below 21 days from') still carries the shortage; the first below-lead-time day is the same either way. Presentation + stock-zero logic.

- **Web page header = cropped banner only** (user 22-Nov-2026 "delete this, I want header like this"): 5_Web + Shortage_Board show the banner image cropped below "Powering America with" (`.banw`), NO separate logo block (`MDU_LOGO_B64` deleted). Presentation only.

- **Header = text in the banner style, no banner image** (user 22-Nov-2026 "this text designed like above, then delete above"): white band, page title in a blue-green gradient pill (white bold), green sub line; banner image and `MDU_BANNER_B64` deleted (both web pages). **5_Web: summary line + "within 21 days" panel moved to a new bottom-bar tab "Summary" next to All** (All / Summary / Runs out / Low cover); Shortage_Board keeps its panel. Presentation only.

- **DAYS OF COVER NAMES** (user 22-Nov-2026 "Name like this: for all same calculation"): one name per calculation on every sheet and web page: 'Days of cover (stock only)' (was Inv. Days / Inv. Days today), 'Days of cover (+ next arrival only)' (was Inv. Days + 21 days arrival), 'Days of cover (+ all in transit)' (was Inv. Days incl. in transit), 'Days of cover (+ next arrival only)' (web pages). NOT renamed (other calculations): MRP Daily / Weekly projection rows 'Inv. Days (stock + in transit ...)', Action List 'Inv. Days at line daily target / plan ...'. Number-format / red-highlight checks accept both prefixes. Names only, no numbers changed.

- **5_Web Summary panel = LAYOUT C** (user 22-Nov-2026 "new labels but this style design"): per material a number row (name + group, 3 'Days of cover' columns) and a detail line under it (stock-zero periods in red, 'next: <date> · qty (+N more, total in 21 d)'). It lived only on branch epic-goldberg-d2dw2a and was never merged; now cherry-picked with the new names. Presentation only.

- **NEXT ARRIVAL ONLY replaces '+ arrivals next 21 d'** (user 23-Nov-2026 "Calc next arrival only from now on"): the Shortage_Board / 5_Web column, scoreboard and Summary panel (`dw`) now = `days_nx` = (Estimated stock today + first arrival, late lines on the base date) / Usage/Day; label 'Days of cover (+ next arrival only)', 'Cover + next arrival:', '<qty> next arrival'. The 21-day arrival sum is gone from those pages. Presentation only (no Excel / calculation change).

- **WEEK BOX / DAILY COVER = HOW MANY DAYS THE STOCK LASTS FROM THAT DAY** (user 23-Nov-2026 "make the week box mean how many days the stock lasts from that day"; RTV showed 357 / 282 on 25-Dec / 01-Jan because the old rule was that morning's stock / that SINGLE day's demand - a rule copied from the planner's WSA column V, never decided by the user for the boxes): `bom_day_cover` now uses the stock that morning (incl. that day's arrival, floored at 0) and uses it up day by day on the PLANNED demand of the next days (holidays / ramps count, zero-demand days count as days that pass, no later arrivals); last day fractional; after the LAST DAY WITH PLANNED DEMAND the rest of the stock runs at Usage/Day (`r['need']`; empty days after the plan end do NOT count as days that last - output-checker round 1: 431 edge protector read 162, 356 front glass 107; now 82 / 21). 'na' = no demand that day. Below 1 = same as before (day's demand not covered), so stock-zero days and the red < 7 rule are unchanged. One source for week strips on all web pages, BOM Weekly / Daily (Excel), 4_Status, risk order, Arrival tab. NOT changed (planner's own WSA formulas): 2_WSA.xlsx Plant Days / Plant + Port Days / column V stay single-day (stock / that day's demand), so on a sheet WITHOUT an ALERT row the 2_WSA.html week cell is still column V. Ask before changing those.

- **TRACING ORDER = SUPPLIER + MATERIAL DESCRIPTION FIRST** (user 23-Nov-2026: "I set rule always read suppliers and material description first"; do NOT ask the planner to correct PPC Final RM Code): a PPC Final RM Code that does not fit the supplier + description is overridden by the traced code and marked CHECK (23-Nov: FLAT invoices 81114864 / 81114865 are 2272 x 1128 "with hole" = 268, PPC says 333). Only exception: identical copied description with different unit prices (unit price + Final RM Code decide, user 21-Nov).

- **WEB PAGES LIST = WITHIN THE MATERIAL'S TRANSIT LEAD TIME** (user 23-Nov-2026 "falls within 21 days: change this rule to within transit lead time"): on all 4 web pages (Material_Shortage, Shortage_Board, 5_Web, 2_WSA) a material is listed when Days of cover (stock only) <= ITS OWN transit lead time (`in_lead_time(r)`, `r['lt']`), not <= 21; the arrival window of the panel / 'Next PPC arrival within the transit lead time' = base date .. base date + its lead time (`a21` keeps its name, `[:lt]`); texts say 'within the transit lead time'. `DASH_HTML_MAX_DAYS` is deleted (no 21 anywhere on the pages except the colour bands 7-21 / 21+, USER 22-Nov). Presentation + selection; no Excel output changes.

- **WEB LIST ALSO = ORDER-NOW STOCK-OUT** (user 23-Nov-2026 "want to see" pallet 535, 18.7 days vs lead time 14, runs out 09-Dec): `in_lead_time(r)` lists a material when Days of cover (stock only) <= its transit lead time OR its first stock-zero day (`k0`) <= lead time + `RELEASE_WINDOW` (7, the ALERT 'ORDER NOW' window, USER) - no new number. Adds pallet 535 and tape RMTTT00000018 (46.8 d vs 45, out 05-Jan). 2_WSA.html rows carry the flag `inl`. Sheets without an ALERT row have `a21=[]` (not listed). Presentation + selection only.

- **ONLY ONE WEB PAGE: 5_Web** (user 23-Nov-2026 "Keep material daily update web page only and delete all script build other web page"): `write_html_dashboard` (Material_Shortage), `write_shortage_board` (Shortage_Board), `write_wsa_html` (2_WSA_<date>.html), the shared `_WEB_CSS` / `_WEB_JS` / `_WEB_TOP` / `_WEB_SRCH`, `DASH_HTML_WEEKS`, `WSA_HTML_WEEKS` and their check_rules entries are DELETED - do not rebuild them. `wsa_html_sheets` became `wsa_open_deliveries` (only the open WSA deliveries for the Details table, filled into `WSA_DLV` in step 2; `wsa_alert_match` still maps sheet -> ALERT row). The writer deletes older Material_Shortage / Material_Dashboard / Shortage_Board / 2_WSA_<date>.html files (exact names, `.html` only - never the 2_WSA .xlsx). Every rule above that names those pages now applies to 5_Web only (list rule `in_lead_time`, week strip, panel, Details, names, colours). Colour rule check in `check_rules.py` now reads the 5_Web code.

- **NO HIDDEN SHEETS BUILT** (user 23-Nov-2026 "Delete script for all sheet that hidden in output"): 4_Status has exactly 2 sheets (BOM Material Dashboard - Weekly + RM Issues). The script that built the 3 hidden ones - 'BOM Material Dashboard - Daily' (`_bom_daily_sheet`, `DAILY_NAME`, `HEAT_ROWS`, its Definitions row), 'RM Material Dashboard' + 'RM Material Dashboard - MIS' (`_dark_rm_dashboard`, `DASH_MIS`, `_wdt_close`, `_WDT_CACHE`, the `use_v` / `cache` parts of `_wsa_dashboard_to`) - is DELETED; do not rebuild it. 1_MRP / Inbound have no hidden sheets. 2_WSA.xlsx keeps the 46 hidden sheets of the PLANNER's own file (carried over, not built by the script; your file is never changed). Not touched: sheets built in 4_Status and deleted again before saving (Materials Status, ALERT, Shortage Watch - they feed the Weekly dashboard).

- **5_Web 'Calendar' tab** (user 23-Nov-2026 "Create a web page calendars show stock zero"; built inside the ONE web page, not a new page): bottom-bar tab Calendar = one grid of the NEXT 4 WEEKS (`CAL_DAYS`=28, user 23-Nov-2026 'next 4 weeks only') from TODAY (user 23-Nov-2026 '4 weeks start from today'; `M.to` = today - base date, so late lines on the base date are not shown); a red day = at least one listed material at stock zero, NO counts (user 'I want material not counting'): the material names are listed by day under the grid, with the PPC arrivals of that day (green triangle on the day, '+qty UOM' per material; user 23-Nov-2026 'add arrivals by day') ( `zd` = days where `bom_day_cover` < 1, never-below-0 rule, same as `runs`); tap a day = only that day; tap a name = match screen; search box filters. Each material's Overview has its own 'Stock-zero calendar'. Presentation only (`calHtml`, `calDraw`).

- **5_Web 'Weeks' tab** (user 23-Nov-2026 "Build a new calendar with week view, next 16 weeks"; second calendar next to 'Calendar', same ONE web page): bottom-bar tab Weeks = 16 week cards (`WEEK_VIEW_WEEKS`=16, Monday of THIS week onward, days before today faded), each with 7 day boxes (red = a listed material at stock zero, amber triangle = PPC arrival, no counts); tap a week to open its materials (stock-zero runs listed one by one, e.g. "23 Nov – 25 Nov, 28 Nov", + arrivals with dates), week 1 open by default; tap a name = match screen; search box filters. Data: `zd` / `aw` per row cover DET_DAYS + 16 weeks. Presentation only (`wkDraw`).

- **WEB PAGE ALWAYS SHOWS THE BOM MATERIAL DASHBOARD** (user 25-Nov-2026: "all info I want to see in BOM material dashboard must be seen in web page always"; standing rule): 5_Web Overview has a card 'BOM Material Dashboard' = every info column of that material's row on the DMS 'BOM Material Dashboard - Weekly' sheet, read from the FINISHED sheet (`_dms_info_capture` after `_bom_weekly_keep` -> `DMS_INFO[(code, item name)]` -> web row `bd`), so a column added to `DMS_WEEKLY_KEEP` appears on the web page with no other change; empty = '-'. The 16 week columns = the Weeks tab (same `bom_day_cover`). Console line '5_Web: BOM Material Dashboard columns shown for N of M'. List rule unchanged (`in_lead_time`). Presentation only.

- **SUBSTITUTE COVER ON THE HEAT MAP + EVERY WEB TAB** (user 25-Nov-2026: "if material is covered by substitute, heat map at the right also reflected" and "EPE 420 covered by EPE 480 ... must mention in all other tabs so that I could know the arrival of substitutes and how the substitutes covering"): `sub_cover` also stores `cv_sub` on the main row (daily cover INCL. the substitute, same meaning as `bom_day_cover`: that morning's own stock first, then the substitute's stock left after its own plan, Use from / until respected, no later arrivals, tail at Usage/Day), `sub_zero`, `sub_x` (substitute arrivals + daily use) and `sub_give` on the substitute row. BOM Material Dashboard - Weekly week cells of the main material = `cv_sub` (italic bold where it adds days; DMS cell comment 'Incl. substitute ...: X days. Own stock only: Y days.', `HEAT_SUBNOTE` -> `_dms_info_capture`). 25-Nov: EPE 420 weeks 8/5/0/0 -> 15/18/11/6. NOT changed: ALERT class / dates, 'Days of cover (stock only)', 'First stock-zero day', 'Stock zero from -> to', risk order (each code's own). Web page: list card + Overview + Arrival strip 'With substitute', 'Substitute' card (Overview, Weeks), Arrival events (substitute arrivals, 'Covered by', 'Stock-out even with', 'Own stock-out'), Weeks / Calendar (blue = own stock zero covered by the substitute; 'covered by' / 'used for' lines), header events.
- **WEB PAGE = ALL MATERIALS, BOM DASHBOARD ORDER** (user 25-Nov-2026: "I want web dashboard show all materials like bom dashboard output"; supersedes the 23-Nov list rule for the list): 5_Web lists every row of the DMS BOM Weekly sheet in its row order (`DMS_ORDER`). `in_lead_time` (flag `il`) still drives the Summary panel, the 'Low cover' filter and the 'N need action' count.
- **PO ETA + PR ETA ON THE WEB PAGE** (user 25-Nov-2026 "I also need PO ETA and PR ETA info. What I can find in BOM dashboard, web page must be also"): `_prpo_by(codes, kind)` = ONE source for the DMS 'ETA - PO' / 'ETA - PR' text (`pr_po_text`) and the web row `pp` ([day offset, 'PO'/'PR', qty]) + `pu` (PO qty without a date). Arrival tab: block with the BOM dashboard columns Arrival at plant / ETA - PO / ETA - PR / Pending GRN, then dated 'PO ETA +qty' (purple) / 'PR ETA +qty' (grey) events, 'PO, no delivery date'; Weeks tab lines 'PO ETA ... · PR ETA ...'. Information only - in no cover / stock-zero number (same as the dashboard).
- **DAYS OF COVER PER PO / PR ETA** (user 25-Nov-2026 "web dashboard also can tell how many coverage days with each PR ETA and PO ETA"): web row `ppc` (per `pp` line: cover on its ETA day, next production day if no demand) + `wkp` (week strip) = `bom_day_cover` on a scenario S / Tp with EVERY dated open PO + PR line added on its ETA (date passed = base date; undated PO qty not in it). Because the daily cover counts no later arrivals, the number on an ETA day = PPC arrivals + all PO / PR up to that day. Shown: Arrival events 'Days of cover if it arrives N (with PPC arrivals + earlier PO / PR)', strip 'If the open PO + PR lines arrive on their ETA' (Arrival tab + Overview Coverage). Information only - in no Excel output, class or stock-zero date. 25-Nov Barcode 2x0.5: own stock-out 09-Jan; PO 07-Jan +150,000 -> 8 days, 14-Jan 8.1, 21-Jan 14.2, 28-Jan 20.3, 01-Mar 23.3.
- **OVERVIEW + ARRIVAL STRIPS = NEXT 16 WEEKS** (user 25-Nov-2026 "want overview and arrival show next 16 weeks as well"): 5_Web Overview 'Coverage' and the Arrival tab use `w16` = `WEEK_VIEW_WEEKS` (16) week boxes, scroll sideways (`.w16s`, 40 px boxes; page width unchanged); the 'With substitute' / 'If the open PO + PR lines arrive' strips there are 16 weeks too (`wk16`, `wa16`, `wks16`, `wkp16`, same `bom_day_cover`). List cards keep the shorter strip (NW = longest lead time of the listed rows). Presentation only.
- **ARRIVAL AT PLANT = DAYS OF COVER, NOT QUANTITY** (user 25-Nov-2026 "web dashboard and BOM material dashboard, Arrival at plant (Final AtA), replace arrival quantity by cover days of that quantity"; since 26-Nov-2026 in EVERY output via `ata_show`): `ata_days(txt, need)` turns every PPC In_Transit quantity into qty / Usage/Day ('Sun 29-Nov: 93,480' -> 'Sun 29-Nov: 16.7 days'; 'Passed, not received' and '+N more' too); PO / PR ETA texts the same (`eta_days`). Need = the row's Usage/Day (ALERT `r['need']`, Action List `pl_need`). Web page shows 'PR ETA' on its own line after ';'; the Excel cell stays one line. No Usage/Day (no demand) = quantities kept. The 'ALERT Arrival lines' sheet keeps the quantities (formula check).
- **PO / PR ETA = DAYS OF COVER + WRAPPED TABLE A..S** (user 25-Nov-2026, picture: "table design from column A to S inspired by Material Status 'Arrival at plant', arrange the info as my picture; also PR and PO qty make it cover day"): DMS BOM Weekly 'ETA - PO' / 'ETA - PR' and the fallback 'PO ETA: ... ;' / 'PR ETA: ...' inside 'Arrival at plant' show qty / Usage/Day (`eta_days`, same rule as `ata_days`; '10-Jan 400,000' -> '10-Jan 37.8 days'; no Usage/Day = quantity kept; `_eta_cols_days`, web page 'BOM Material Dashboard' card copies the text). PR ETA follows PO ETA on the SAME line ('  ;  '; the web `brk` breaks it). Information only; ALERT / email / 1_MRP unchanged. Checker not run (presentation + information columns, no calculation).
- **BOM WEEKLY LOOK = 5_Web page** (user 27-Nov-2026: "design sheet weekly in 4_Status like web page summary style"; REPLACES the 1_MRP 'Material Status' look of 25-Nov-2026): `_dms_weekly_style`: white page, no grid, light grey row lines only, band row 1 = one coloured title per group (MATERIAL ink / STORE blue / COVER green / SHORTAGE red / WAREHOUSE purple) + blue->green gradient title over the week columns, header row 2 muted grey bold 9 pt with a group-colour underline (only wrapped row, 45 high), Segoe UI, ONE line per cell, EVERY data row 19 high. Material name bold; 'Days of cover' bold, red under `COV_RED_DAYS`; First stock-zero day / Stock zero from -> to red; PO ETA purple; substitute blue; week cells = web chips (old red / amber / green -> `C9372C` / `F6C4C4` / `E9EDE9`, same bins, white gap). Group band rows keep fill E8ECF1 (`_dms_cover_filter` finds them by it). Do not wrap data rows again. Presentation only.
- **A4 LANDSCAPE PRINT** (user 25-Nov-2026 "print A4 landscape each sheet BOM material dashboard, alert sheet and action sheet. Find best format"): `print_a4_all` (called in `_safe_save`) -> `print_a4` on 1_MRP 'ALERT' + 'Action List' and DMS 'BOM Material Dashboard - Weekly': A4 landscape, page cuts ONLY between column groups (band-row merges; Action List `PRINT_AL_GROUPS`; BOM Weekly WAREHOUSE cut before 'Pending GRN Qty'), code / name columns + band and header rows repeat on every page, zoom = largest that gives the fewest pages wide (floor `PRINT_MIN_SCALE`, 8 % width safety), pages across then down (first sheets = highest-risk materials complete), header sheet name + date, footer file + 'Page x of y', no gridlines. 25-Nov: ALERT 2 wide x 3 tall at 58 %, Action List 3 x 5 at 60 %, BOM Weekly 3 x 2 at 57 %. Print settings only.
- **DMS BOM WEEKLY: E..S WIDTH 11 + FILTER < 22 DAYS** (user 25-Nov-2026 "column width 11 from column E to S. pdf, I want to print Days of cover (stock only) less than 22 days only"): end of `_dms_weekly_style` sets columns E..S to width 11 (long texts cut at the width; full text in the cell / web page) and `_dms_cover_filter` puts the autofilter 'Days of cover (stock only)' < `PRINT_MAX_COVER` (22) on the sheet and hides the other material rows ('-' / no demand hidden too); a group band row stays when one of its materials shows; legend rows under the table stay. Excel 'Clear filter' shows all. 25-Nov: 19 materials, 2 pages (59 %). Values unchanged; the web page still lists every material.
- **BOM WEEKLY PRINT = ALL ROWS ON ONE PAGE HEIGHT, COLUMNS ACROSS PAGES** (user 25-Nov-2026 "pdf fit rows in page, I don't want to fit everything in 1 page"; replaces the 3-pages-wide cut AND the one-page try of the same day, for this sheet only): `print_rows_fit` (from `print_a4_all`) = A4 landscape, zoom = largest (max 100 %) at which every shown row fits the page height (Excel 'Fit all rows on one page'; 91 % on 25-Nov), columns cut between band groups (a group wider than a page is cut by column), Item name repeats. ALERT / Action List keep `print_a4`. Print settings only.
- **'Covered by substitute' TEXT** (column renamed from 'Covered by substitute (code, when, qty)', user 25-Nov-2026; text order = the user's own sample): `sub_cover` text = '11.0 days of cover (stock only): own 4.4 + EPE 480 6.6.' (user sample 25-Nov-2026, total first) > 'Then <sub> <code> covers dd-Mmm → dd-Mmm qty UOM; ... (total ..; its stock .. + arrivals ..).' > 'With <sub>: stock zero ...'. NO 'own stock + own arrivals cover to' sentence. Text only; no number changed (also ALERT note / email / web page `subcov`).
- **POD + In transit STATUS** (user 25-Nov-2026 "bom material dashboard and web dashboard has info about POD and STATUS column from in transit"): PPC In_Transit columns POD / STATUS of the lines counted as in transit for the material (same lines + WSA-sheet / supplier split as 'Pending GRN Qty'; `read_ppc_in_transit` keeps `pod`, `it_pod_status` / `pod_text` / `status_text`). DMS BOM Weekly columns `POD_HEAD` 'POD (port of discharge)' + `STS_HEAD` 'In transit STATUS' right after 'Pending GRN Qty' (WAREHOUSE group, in `DMS_WEEKLY_KEEP`; so also on the web page 'BOM Material Dashboard' card and the Arrival-tab block). POD = one port: its name; 2+ ports 'PORT: dates'; STATUS = 'status text: arrival dates' per distinct status (4 dates + '+N more'; ' | ' inside the PPC text -> '; '); '-' = nothing in transit. Columns E..S stay width 11 (user rule), so STATUS is cut on screen / print (full text in the cell and on the web page). Information only - in no calculation; ALERT / email / 1_MRP unchanged.

- **CALENDAR / WEEKS RED = RM ONLY** (user 25-Nov-2026 "Web, calendar and week tab, only red when rm material is at stock zero"): 5_Web Calendar + Weeks tabs colour a day red (or blue = covered by substitute) only for a material in group RM (`r.g=='RM'`, `_heat_grp`); PM / CS stock-zero days are still listed under the day / week (grey square, grey text), the day stays uncoloured but tappable. A single material's own calendar (Overview) is unchanged. 25-Nov: 29 + 30 Nov no longer red (only PM / CS at zero). Presentation only.
- **1_MRP SHEETS 1-5 LOOK = 5_Web Summary tab** (user 27-Nov-2026: "design sheet 1 to sheet 5 of 1_MRP like 5_Web summary tab style"): `web_style_mrp(path,specs)` (src/08, called in `build_new_mrp` right before `post_flush`; specs = header row, band row, title rows per sheet) uses the `_dms_weekly_style` palette: Segoe UI, white page, no grid, light grey row lines, one coloured group band (STORE blue / PRODUCTION green / WAREHOUSE purple / TRANSIT LEAD TIME red / BELOW + INVENTORY orange; dark title = blue->green gradient), muted grey 9 pt header with group-colour underline, old status fills -> web chips (light red / grey), old font colours -> web families. Row heights are kept (wrapped lines stay visible, min 19). Values, formulas, columns, merges, filters, print setup unchanged (checked: 0 value diffs on all 25 sheets). The 4_Status copy of Material Status is made earlier and is not touched. Presentation only.

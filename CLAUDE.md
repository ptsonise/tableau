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
| ALERT list = under 21 days of stock or a stock-zero day inside the lead time (`alert_eval` + `build_alert_rows`, built ONCE) | Inbound first tab 'ALERT' (`add_alert_inbound`, the same rows) | first tab 'ALERT' (`write_alert_sheet`) |

Nov 2026 check: 280 shipments, 7,056,776 in transit (MRP units) on 20 materials - identical per material in both outputs.

Run: `python update_inbound.py` (Python 3.9+, pandas, openpyxl, lxml). A run takes about 45 s: every run also writes the
aligned WSA file (see `--align-wsa` below; `--no-align` = old behaviour, about 20 s).

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
  Open question (Nov 2026): a 0 in the newest date column (e.g. pine wood pallets PMPKM00000447 / 535 on 05-Nov) is
  taken as a real count of 0 - ask the user before changing this.
- **OEM** (customers with OEM = Yes on Settings): their plan and the **JW stock at the OEM** are left out of every projection,
  shortage date and risk list. Only the Action List "with OEM" columns add the JW stock back.
  In code: the OEM step in `build_new_mrp` subtracts `st_jw` from `b['stock']` once and sets `b['jw_out']=True`.
  Any code that runs after it must **not** subtract `st_jw` again (check `jw_out`). This was the Nov-2026 Stock Check bug.
- **Substitute groups** (Settings SUBSTITUTES): projected as one pool (group stock + in transit vs group demand).
- **All sheets of MRP_Daily are visible** (user's choice, Nov 2026). The self-check keeps skipping Check / Stock Data /
  No Demand / Substitute Advice (not in shortage order).
- **Daily targets and OEM = Action List ONLY** (user rule, Nov 2026). Columns 'Production Daily Target_with/without OEM',
  'Inv. Days at line daily target...', 'to OEM Stock' stay there. **Every other sheet follows the module plan**:
  `x['pl_need']` = average plan per production day, next 16 weeks (OEM plan left out, substitute group pooled);
  `x['pl_days']` = (own stock + in transit) / pl_need. Used by Stock Management (Need/day, reorder level), Risk Charts,
  Cover Heatmap, risk class 'Under 21 d', risk order, Substitute Advice, the daily report, No Demand, MRP by Line
  ('This line: daily plan need'; the Action List target / OEM columns are not copied there).
  Action List shows BOTH scenarios (user rule, Nov 2026 - do not ask again):
  - daily TARGET: 'Below 21 Days From' / 'Days Late' use the target need (`x['cw']`, with OEM target), next to the target columns;
  - MODULE PLAN (gold band): 'Daily plan need (module plan)', 'Inv. Days at module plan (stock + in transit)',
    'Below 21 Days From (module plan)', 'Days Late (module plan)' (`x['cw_p']` = Risk Charts values). Copied to MRP by Line.
  - Status follows the plan. Stock-zero periods on the Action List = 'Short from → to (with in transit)' only (all periods,
    ALERT label, lead-time slack). 'Short From (Stock = 0)' was removed from the Action List (duplicate, user OK Nov 2026);
    it stays on Risk Charts / RM Issues.
- **MRP Daily / Weekly layout** (Nov 2026): 'Firm receipt (Open PO, no vessel)' and 'Days of cover (stock + in transit + Open PO)'
  rows exist only for materials with an Open PO (rowmap `ro` / `rj` = None otherwise - guard every use). Supplier column shows the
  single real supplier (e.g. FLAT) instead of 'All suppliers' (`b['_sup1']`; row keys unchanged). Supplier blocks only when 2+ suppliers
  supply the material (`multi_sup`). Days of cover = projected stock at day end / average plan usage of the next COVER_AVG_DAYS (21).
- Material status: CRITICAL / AT RISK / WATCH / OK vs each material's own target days (`material_status()`); line status uses COVER_* (`line_status()`).
- **ALERT sheet = first tab of MRP_Daily AND of the Inbound file, the same rows** (user request Nov 2026: "list of less than 21 days
  inventory, how many days stock zero within lead time and zero from-to ... many sheets have this info but too complicated - one sheet,
  simple and strong alert"). `build_alert_rows(A,d0,nd)` (end of `build_new_mrp`, stored in `ALERT_ROWS` / `ALERT_D0`) -> `alert_eval`:
  days of stock = stock today / `pl_need`; days incl. in transit = `pl_days`; stock-zero day = S < 0 (the Action List projection: stock +
  MIS in transit arrived - plan; the day's need is not covered; on a WSA sheet = Plant + Port Days (PPC In Transit) below 1 - its red
  cell (< 0) comes a day later); zero days inside the lead time = S < 0 days from today to today + `lt_used` - 1; listed when days of
  stock < REPORT_DAYS (21) or zero days inside the lead time > 0. Classes: STOCK-OUT IN LEAD TIME (red, EXPEDITE / alternative),
  STOCK-OUT LATER (orange, ORDER NOW / ORDER BY first zero day - lead time, RELEASE_WINDOW like `risk_priority`), UNDER 21 DAYS
  (yellow, in transit covers). Columns: Alert, Code, Material, UOM, Stock today, Need / day, Days of stock, Days incl. in transit, Lead
  time, Stock-zero days inside the lead time, First stock-zero day, Stock zero from -> to (every period; '⚠' = starts inside the lead
  time), Next arrival, Action; three coloured counters + one definitions line on top. Substitute group = one row (`grp_label`, pooled;
  followers skipped). A material on 2+ WSA sheets (333) = one row per sheet from the aligned sheet's own numbers (`WSA_REFS[0]` groups:
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
  references in other sheets / defined names follow, `_ref_fix`; the Dashboard row reads the front sheet; back sheets have no Dashboard
  row). A one-glass block whose BL# lines show the partner glass too (Glass G12R Topcon Subanci = 356 + 333) is split the same way:
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
  2,931 / 430,189 (same shipments, by supplier instead of by sheet). Open point told to the user: once the FLAT HK glass is in stock,
  MB52 only says FLAT -> it counts on Subanci Back; adding Encap rows to ALLOWED VENDORS (FLAT first) shares FLAT stock by demand.
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
  Pine wood pallets show stock 0 from the 05-Nov count of 0 (open question above).
- **Inbound 'Below zero'** (Nov 2026 user request): `below_zero_periods()` = from the sheet stock date, periods where the sheet's
  Plant Days < 0 and where Plant + Port Days < 0 ('na' days keep the previous state). Summary columns R/S ('... < 0 (from -> to)',
  `periods_text`) + sheet 'Below zero' (one row per period, with the next delivery on/after From).
- Known: EVA 420 / EPE 420 sheets say 'Pcs' but the numbers are M2. Cell Fixation tape: PPC BOM has RMTTT00000018 (100 m)
  0.015-0.018/module -> RMTTT00000010 (50 m) x2 = 0.03-0.036; the old sheet used 0.008 (asked the user to check).

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

## Working with the user
- Correct the grammar and vocabulary in the user's message first (US English), then restate the request as a developer requirement, then answer.
- **Two daily outputs, one logic** (see the top of this file): a rule asked for the WSA daily goes into MRP_Daily in the same version
  (and the other way round); prove both outputs in the comparison.
- **After every change to `update_inbound.py`, send the updated file to the user (attachment) in the same reply**, so they can run it.
  Send it only after the change is tested, and say in one line what changed in that version.
- **Also send the output files** (`MRP_Daily_<date>.xlsx` and `Inbound_<date>.xlsx`) from a fresh run of that script on the
  user's latest uploaded inputs. Put them in `output/` (git-ignored - company data, never commit or push them).
- Use simple, clear language. Planning terms (MRP, MB52, JW, OEM, BL, ETA) are fine.

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

Outputs: `MRP_Daily_<date>.xlsx` (Action List, MRP Daily, Stock Check, Check, ...) and `Inbound_<date>.xlsx` (deliveries, To chase, By Week).

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

- **WSA alignment = DEFAULT on every run** (Nov 2026: the user runs plain `python update_inbound.py` and expected the WSA file;
  `--no-align` skips it, `--align-wsa` is still accepted). Step 1 reads a temp COPY of the WSA file (the output may replace
  the file it ran on; Windows cannot replace an open file) with console output muted (`_QUIET`, errors still shown); an
  alignment error is reported and step 2 then runs on the user's file as it is. The console ends with 'WSA ALIGNED FILE SAVED: <path>'.
  Aligned sheets open at today's row (sheetView / pane topLeftCell) and today's Balance takes the style of the user's last
  typed balance (yellow); later Balance rows that had a typed style get the formula style. The user once thought the file was
  unchanged because the sheets opened at row 1 (2024 rows).
- **`--align-wsa`** (Nov 2026, user request: "recalculating from today and live excel"): rebuilds the WSA planning columns so
  they match MRP_Daily, as LIVE Excel formulas. Code: `align_capture()` (called in `build_new_mrp` right after the OEM step),
  `align_wsa()` (writes the file), main block (step 1 = calculation run to a temp folder, step 2 = normal run on the aligned
  file via subprocess).
  - Output: `<WSA name> (aligned <date>).xlsx` next to the WSA file; the original is never changed.
  - New sheet `PPC_Plan_Req` (appended last, so sheet indexes stay): one row per date x line x customer x planning block:
    Modules (pp_rows), Qty per module (= the block's codes' `cust_line` use / modules, so per-day exact), Requirement `=H*I`,
    'Counted in Balance' (No = OEM), Qty source (PPC BOM / MIS BOM from `plan_basis`). Rows before today are kept on a re-run
    (past rows of the planning sheets read them).
  - Rewritten: every visible planning block, rows dated today or later, columns from Balance up to (not including) Delivered.
    Customer columns = SUMIFS(Modules, block, column header, date); Planned = SUMIFS(Modules, block, Counted=Yes, date);
    Demand = SUMIFS(Requirement, block, Counted=Yes, date); Balance today = typed own stock (MB52 / PPC count rolled forward,
    JW out) so the script's stock date becomes today; later rows = prev Balance - prev Demand + prev Delivered.
    Delivered and everything right of it, and all past rows, are untouched (user: "data I can't control").
  - Customer column labels come from the WSA 'Lookup' sheet (plan line -> MIS plant: 'USA Line 3' -> 'L3' ...).
    `WSA_OEM_IN_BALANCE=True` would count JW stock + OEM plan.
  - Edited at XML level (pivots, Power Query, tables, styles kept): shared formulas expanded, calcChain removed, fullCalcOnLoad
    set; every formula gets its cached value (SUMIFS mirrored in Python, MW / Plant days by `_xl_eval`) so the script reads them.
  - Proof used (Nov 2026): 0 cells changed outside the rebuilt area; all 48,130 rebuilt cells equal LibreOffice's own
    recalculation (wrong cached values planted in a test copy were corrected by LibreOffice, so it really recalculates);
    a simulated next-day re-run kept the previous day's rows; without the flag the outputs are unchanged.
  - Known: EVA 420 / EPE 420 sheets say 'Pcs' but the numbers are M2. Cell Fixation tape: PPC BOM has RMTTT00000018 (100 m)
    0.015-0.018/module -> RMTTT00000010 (50 m) x2 = 0.03-0.036; the old sheet used 0.008 (asked the user to check).

## Known data quirks (NOT script bugs)
- WSA sheet `Glass G12R Topcon Subanci` lists only RMGLS00000356 but its Demand counts front + back glass. MIS splits it
  into 356 (front) + 333 (back). The Check sheet "Plan" mismatch for glass is expected; the totals match.
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
- **After every change to `update_inbound.py`, send the updated file to the user (attachment) in the same reply**, so they can run it.
  Send it only after the change is tested, and say in one line what changed in that version.
- **Also send the output files** (`MRP_Daily_<date>.xlsx` and `Inbound_<date>.xlsx`) from a fresh run of that script on the
  user's latest uploaded inputs. Put them in `output/` (git-ignored - company data, never commit or push them).
- Use simple, clear language. Planning terms (MRP, MB52, JW, OEM, BL, ETA) are fine.

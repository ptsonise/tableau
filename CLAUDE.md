# CLAUDE.md

Short rules for the raw-material planner (USA solar module plant). Code = `src/*.py` (15 files, loaded by `update_inbound.py`). Loaded every session - keep it short.
**Full detail per feature (function names, examples, history): `docs/RULES_DETAIL.md`** - read the section you need before
changing that area. New rule = one line here + detail there. Old history: `git log -p`.

## Code map (READ ONLY THE FILE YOU NEED - never `update_inbound.py` / `dist/`; find a function: `grep -n "def name" src/*.py`)
`src/*.py` run in name order in ONE shared namespace (same as the old single file). New code goes in the file that owns the area.
01 setup/config/constants | 02 workbook cache, PPC In_Transit read | 03 daily stock, MIS rebuild | 04 inventory, in transit, substitute advice |
05 Check sheets, CS/PM | 06 user BOM, calc sheets | 07 ALERT | 08 RM issues, stock dashboard | 09 Material_Daily_Update html, DMS |
10 DMS weekly, GR / PO readers, count history | 11 Settings, supplier map, module plan | 12 build_new_mrp | 13 map In_Transit, glass |
14 WSA align (WSA_Daily) | 15 main run. `python build_single.py` -> `dist/update_inbound.py` = ONE file for the planner's PC (git-ignored).

## Top rules (never break)
1. **The planner's WSA file is NEVER changed.** WSA_Daily_<date>.xlsx is a new file next to it; outputs carry the run date.
2. **Two daily outputs, one logic**: every calculation / mapping rule goes into BOTH WSA_Daily (+ Inbound) AND MRP_Daily in the same
   change, without being asked; show both in the before/after. Only presentation the user placed on one sheet stays there.
3. **Numbers register** (below): USER numbers never change without asking; `python check_rules.py` must PASS.
4. **Test before saying done**: run before/after on the real inputs, compare every sheet; only expected cells change.
5. **Never commit Excel inputs / outputs** (company data). Outputs go to `output/` (git-ignored).
6. **Delete unused code** (user rule) - no dead functions / constants. Do not rebuild deleted sheets / pages (list in RULES_DETAIL).

## Inputs / outputs / run
Inputs (script folder): `WSA - Raw Material planning per day.xlsx`, `MRP.xlsx` (Settings sheet), `MIS RM Dashboard - WSA.xlsx`,
`PPC*.xlsx` (WSA Plan, OEM Plan, Revised Plan, BOM, Count, GR, In_Transit, MB52, PR / PO, Pending GR, PM_CS_Usage), MB52.xlsx only
when PPC has no MB52 sheet. Outputs: MRP_Daily, Inbound, Daily_Material_Status, WSA_Daily (.xlsx),
Run_Summary, Material_Daily_Update (.html = the ONLY web page). Run `python update_inbound.py` (Python 3.9+, ~60 s);
`--no-align`, `--keep-old`. Old outputs deleted after a full run (exact script names only).

## Calculation rules (detail: RULES_DETAIL 'Key rules')
- **Demand** = module plan (PPC WSA Plan + Revised Plan on top, OEM plan out) x qty per module (PPC BOM, else MIS Standard BOM).
  Plan customer -> MRP customer only via LINKS / Settings MIS LINK.
- **Stock** = MB52 (Unrestricted + QI), replaced by PPC Count (main table, 'Physical Count Qty_BOM' + 'Count date') rolled forward
  (count - usage + PPC GRN receipts since count). BASE DATE = MB52 file date when older than today.
- **PPC GR sheet = reference only** (no calculation). A receipt goes to ITS OWN supplier.
- **OEM**: OEM plan and JW stock at OEM are in NO calculation (stay visible as data). Subtract JW once (`own_stock` before the OEM step).
- **In transit = PPC In_Transit ONLY** (MIS ignored). A line counts only with ETD + ETA + BL# + Final AtA; qty = Final Pending Qty;
  date = Final ATA2 only when Final AtA has a date. PPC dates are the latest; never move the planner's WSA rows unless asked.
- **Material of a shipment**: supplier + description first (PPC Final RM Code overridden + CHECK if it doesn't fit); WSA PO + BL#;
  same description with different unit prices -> unit price + Final RM Code decide.
- **Substitutes = separate materials everywhere** (own stock / plan / transit); substitute cover shown as presentation only (`sub_cover`).
  Still folded: busbar 016/032/005/033, tape roll lengths (CODE_SWAP, BOM code = the material), CS/PM '_Alt' lines.
- **Settings ALLOWED VENDORS in every projection** (`vendor_pools`); customer goes with supplier at the SOURCE of every list;
  DEAD supply (vendor no customer may use, e.g. 312 FLAT) is out from day 1.
- **Front / back glass = two materials** (own WSA sheets). Units stay as SAP / MIS / PPC BOM give them.
- **Every day threshold = the material's transit lead time** (Settings TRANSIT LEAD TIME > median WSA > fallback 49 RM / 14 CS-PM).
- **Usage/Day** = average plan per production day over the lead time. Daily targets / OEM columns = Action List only.
- **Stock never below 0** in every cover / stock-zero calculation (`floor_path`). NOT in the planner's WSA formulas (Balance, Plant
  Days, column V) or MRP Daily/Weekly projection rows - ask before changing.
- **Daily cover / week box** = days the stock lasts from that day (`bom_day_cover`), one source for all strips and dashboards.
- **CONSUMABLES** = user's 7 codes (`CONSUMABLE_CODES`); else PM / RM by PPC Count Item (`item_category` = one rule).
- **Days-of-cover names**: 'Days of cover (stock only)', '(+ next arrival only)', '(+ all in transit)'.

## Output layout rules (detail: RULES_DETAIL 'ALERT', 'Dashboards', 'WSA_Daily')
- **MRP_Daily**: ALERT first tab, then Material Status; ALERT (formulas) / Calc / line sheets; ETD alert; NO dashboard sheets;
  no hidden sheets. No 'Action' column on any sheet. No INVENTORY / QUALITY on ALERT / Material Status.
- **Inbound**: no ALERT sheets (only 'ETD alert').
- **Daily_Material_Status** = exactly 2 sheets: 'BOM Material Dashboard - Weekly' (columns = `DMS_WEEKLY_KEEP`, Material Status look,
  one line per data row, 19 high, E..S width 11, filter Days of cover < 22) + 'RM Issues'.
- **WSA_Daily.xlsx**: one layout per sheet (rows 1-6 info, header row 7, columns A..AL), past rows = values, today on = formulas with
  cached values; must read back before replacing.
- **Material_Daily_Update.html** = only web page; light theme; shows every BOM Weekly row and column (`DMS_INFO`), PO / PR ETA,
  Calendar + Weeks tabs (red = RM only). Presentation only.
- **ETA - PO / PR, Pending GRN, POD / STATUS, Arrival at plant as days of cover** = information only, in no calculation.

## Numbers register (enforced by `check_rules.py`)
USER = planner decided: never change without asking. DEFAULT = unconfirmed: don't change silently, ask when it matters.
A number not listed = DEFAULT; add it here + `check_rules.py` when touched. Change code + this table + check_rules.py in ONE commit.

| Number | Value | Owner |
|---|---|---|
| Every day threshold | material's transit lead time | USER |
| `COV_RED_DAYS` | 7 | USER |
| Web colours / weeks | red <7, light red 7-21, grey 21+ / 16 weeks | USER |
| `BOARD_MAX_DAYS` | 63 | DEFAULT |
| `CONSUMABLE_CODES` | 7 codes | USER |
| `LEAD_TIME_RM` / `LEAD_TIME_CSPM` fallback | 49 / 14 | USER |
| `PLAN_AVG_DAYS` | 49 | USER |
| `CUST_TARGET_DEFAULT` | Pepper 3000, Lucky 7 3000, Crimson 3600, SB 3600, Snowflake 4000 OEM | USER |
| `ALERT_HORIZON` / `RISK_DAYS` / `RELEASE_WINDOW` / `ETD_SOON` | 112 / 5 / 7 / 7 | USER |
| `RESCHED_OUT_X` / `SHORT_HORIZON` / `ATA_SHOW` / `MOVE_DAYS` | 3 / 60 / 4 / 14 | USER |
| `SHORT_LEAD` / `LONG_LEAD` | 20 / 75 | USER |
| `DASH_STEP2` / `DASH_BLUE1` / `DASH_BLUE2` / `DASH_MIN_LT` | 14 / 14 / 39 / 14 | USER |
| `PRINT_MAX_COVER` | < 22 days | USER |
| `IT_STATUS_ONLY` / `UNLIMITED_CODES` | False / PMPKM338, 523 | USER |
| PPC line without ATA (no Final AtA column) | Revised ETA Port + 5 | USER |
| PPC In_Transit line needs | ETD + ETA + BL# + Final AtA | USER |
| Status WATCH / line CRITICAL | lead time + 14 / under 2/3 of lead time | USER |
| Tracing / Inbound AQ / In_Transit read stop | 2+ points / 3+ days / 500 empty rows | USER |

## How to work on this repo
1. Get a concrete symptom first (material code, sheet, value shown, value expected). Don't guess.
2. Reproduce with the real input files (uploaded by the user; not in the repo).
3. Prove the fix: before/after run, compare every sheet of both outputs; before/after table for affected materials.
4. Python 3.9+: no same-kind quotes inside f-string expressions. Check:
   `/usr/bin/python3.11 -c "import ast,glob;[ast.parse(open(p).read()) for p in glob.glob('src/*.py')]"` and `python check_rules.py` (PASS).
5. LOGIC changes only (calculation, mapping, projection, date, stock, plan, vendor rule): run the `output-checker` subagent (tell it
   the before-script, e.g. `HEAD~1`), fix proven findings, max 3 rounds; show its verdict. Not for text / label / colour / layout.

## Working with the user
- First correct the user's grammar (US English), restate as a developer requirement, then answer. Simple language.
- After every script change: run `python build_single.py` and send the tested `dist/update_inbound.py` + fresh MRP_Daily / Inbound outputs, one line on what changed.
- Fewer messages: when choices are open, build the recommended option and say so; ask only when the answer changes the numbers.
- Say when a new session is better (context summarized + task done, big new feature, new day's run). Handover: branch -> PR to
  master on request, the 5 input files to upload, open points below.

## Known data quirks (NOT bugs)
MB52.xlsx date = Windows file date. Planner's 'Glass G12R Topcon Subanci' sheet mixes 356 + 333 (aligned file splits it).
268 = back glass, 280 = front. 'Clenera' vs 'Cleanera' Crimson = fix on Settings. "In PPC In_Transit, not in WSA" = planner not
updated yet. DMS 'Other Materials' mostly 'No demand' (PPC Count col K XLOOKUP).

## Open points (ask the user; remove when answered)
- SAP GR posted while PPC In_Transit still pending = counted twice (invoice 81117083). Offered: drop those lines. Not done.
- Speed (~70 s): which outputs are opened daily + 'Timing:' line from user's PC - not answered.
- PPC Count second list at row 170: newer count or scratch? Ignored until answered.
- Low: RM Issues 'Project' blank for EPE 480 + 420; Stock Management KEY RISK counts that group twice.

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

Outputs: `MRP_Daily_<date>.xlsx` (Action List, MRP Daily, Stock Check, Check, ...) and `Inbound_<date>.xlsx` (deliveries, To chase, By Week).

Run: `python update_inbound.py` (Python 3.9+, pandas, openpyxl, lxml). A run takes about 25 s.

## Key rules (how the numbers should behave)
- **Demand = MIS plan** (Total Plan Req x BOM usage), split by customer and line. Not the WSA sheet demand.
- **Stock = your MB52** (+ physical count from Stock Count when present, rolled forward to today).
- **OEM** (customers with OEM = Yes on Settings): their plan and the **JW stock at the OEM** are left out of every projection,
  shortage date and risk list. Only the Action List "with OEM" columns add the JW stock back.
  In code: the OEM step in `build_new_mrp` subtracts `st_jw` from `b['stock']` once and sets `b['jw_out']=True`.
  Any code that runs after it must **not** subtract `st_jw` again (check `jw_out`). This was the Nov-2026 Stock Check bug.
- **Substitute groups** (Settings SUBSTITUTES): projected as one pool (group stock + in transit vs group demand).
- Material status: CRITICAL / AT RISK / WATCH / OK vs each material's own target days (`material_status()`); line status uses COVER_* (`line_status()`).

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
- Use simple, clear language. Planning terms (MRP, MB52, JW, OEM, BL, ETA) are fine.

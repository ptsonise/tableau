# CODE_MAP.md - the 15 code areas of the raw-material planner

The code lives in `src/*.py`. `update_inbound.py` is a small loader: it runs the 15 files in name order as ONE program, so a
daily run is still one command: `python update_inbound.py` (~60 s).
For your PC, `python build_single.py` joins the 15 files into one file: `dist/update_inbound.py` (same code, same output).

**How to ask for a change:** name the area (file) or the sheet you see, the material, the value shown and the value expected.
Find any function: `grep -n "def function_name" src/*.py`.

## Big picture (order of a run)
Read the inputs (01-03) -> classify materials, supply, substitutes (04) -> checks, BOM, calc sheets (05-06) -> ALERT (07) ->
issues and dashboards (08) -> email, web page, Daily_Material_Status (09-10) -> settings and plan (11) -> build MRP_Daily (12) ->
match in-transit shipments (13) -> build WSA_Daily and vendor rules (14) -> main run, saves all files (15).

## The 15 areas

| # | File | Lines | What it does | Key functions | Change it when you want... |
|---|---|---|---|---|---|
| 01 | `01_setup_config` | 293 | Program start: imports, Excel speed patches, DEFAULT numbers and shared lists | `lt_of`, `tgt_days`, `usage_day`, `material_status`, `project_of` | a different default number (most are in the Numbers register: ask first), the status rule (CRITICAL / AT RISK / WATCH / OK) |
| 02 | `02_workbook_cache_ppc_in_transit` | 959 | Opens the input workbooks once (cache), finds the PPC file, reads PPC In_Transit lines, MB52 and Revised Plan; "Other materials" and "Below zero" calculations | `wb_ro`, `find_ppc`, `ppc_it_rows`, `normalize_in_transit`, `read_revised_plan`, `ppc_plan_sheet`, `other_materials_calc`, `below_zero_periods` | a change in how an input file or sheet is found / read |
| 03 | `03_daily_stock_mis` | 1205 | Daily stock sheet from the MIS dashboard, physical counts, A4 print helpers, attaches the WSA data | `prepare_mis`, `read_phys_counts`, `build_daily_stock`, `print_a4`, `print_rows_fit`, `make_daily_stock_xlsx`, `attach_wsa` | stock / count reading, print layout (A4, rows-fit) |
| 04 | `04_inventory_in_transit_substitutes` | 777 | In Transit data and In Transit Status sheet, own stock (JW stock subtracted), days of cover, material groups / PM-RM category, Substitute Advice | `build_mis_intransit`, `own_stock`, `cover_days`, `item_category`, `read_bom_groups`, `make_substitute_advice_xlsx`, `make_intransit_status_xlsx` | in-transit status, substitute advice, how a material is classified |
| 05 | `05_check_sheets_cspm` | 500 | Check sheets (MIS vs WSA vs MRP differences), MIS layout check, CS / PM (consumables, packaging) BOM and stock | `build_checks`, `mis_layout_check`, `make_check_xlsx`, `read_user_inputs`, `module_plan`, `cspm_from_bom`, `cspm_alt_fold` | the Check sheet, CS / PM logic, '_Alt' lines |
| 06 | `06_user_bom_calc_sheets` | 769 | Your BOM (MRP.xlsx), usage per module, Calc sheets (target, plan), movement log, run summary, risk helpers | `load_user_bom_rm`, `bom_usage`, `rm_plan_from_mis`, `make_calc_plan_xlsx`, `make_calc_target_xlsx`, `run_summary`, `risk_info`, `is_rm` | BOM / usage logic, Calc sheets, Run_Summary |
| 07 | `07_alert` | 1173 | The ALERT sheet (rows, why, floor at 0, stock fields), ETD alert, days of cover per day, consumables, cover heatmap | `build_alert_rows`, `alert_eval`, `floor_path`, `write_alert_sheet`, `write_alert_formulas`, `write_etd_late_sheet`, `bom_day_cover`, `is_consumable`, `add_cover_heatmap` | anything on ALERT: columns, colours, risk rows, cover days, ETD alert |
| 08 | `08_rm_issues_stock_dashboard` | 896 | RM Issues log, Stock Check sheet, one workbook style, Stock Management dashboard | `add_rm_issues_sheet`, `load_rmi_prev`, `add_stock_check_sheet`, `unify_workbook_style`, `add_stock_mgmt_dashboard` | RM Issues, Stock Check, fonts / colours of the workbooks |
| 09 | `09_daily_email_html_dms` | 1053 | Daily email (.eml / .html), Material_Daily_Update web page, Daily_Material_Status workbook, Pending GRN, POD status | `write_daily_email`, `write_material_daily_update`, `make_daily_material_status`, `pending_grn`, `pod_text`, `wsa_open_deliveries` | email layout, the web page, Daily_Material_Status |
| 10 | `10_dms_weekly_gr_po_count` | 1324 | BOM Weekly dashboard styling, stock data and movement log sheets, GR / PR / PO readers (information only), count history and roll-forward, pack-size check | `make_stock_data_xlsx`, `read_gr_lines`, `read_pr_po`, `pr_po_text`, `read_count_history`, `estimate_counted_stock`, `apply_count_by_supplier`, `make_pack_check_xlsx` | BOM Weekly look, PO / PR ETA text, count roll-forward, unit and pack checks |
| 11 | `11_settings_supplier_module_plan` | 925 | Settings sheet (reading and writing), substitute links, supplier map, module plan, line view, BOM and Count sheets | `read_settings`, `resolve_substitutes`, `make_settings_sheet`, `read_supplier_map`, `apply_supplier_map`, `module_plan_from_pp`, `module_plan_sheet`, `make_line_view_xlsx` | a NEW setting, supplier / customer map, module plan reading |
| 12 | `12_build_new_mrp` | 960 | One big function, `build_new_mrp`: builds the MRP Daily, MRP Weekly and MRP by Line sheets | `build_new_mrp`, `is_unlimited` | the projection rows of MRP_Daily / MRP Weekly (ask before changing the floor at 0) |
| 13 | `13_map_in_transit_glass` | 573 | Matches each PPC In_Transit shipment to YOUR material (supplier + description, PO + BL#, unit price), reads PPC In_Transit arrivals, captures WSA numbers | `map_in_transit`, `in_transit_catalog`, `wsa_find`, `wsa_ref_index`, `read_ppc_in_transit`, `ppc_arrivals`, `align_capture` | which material a shipment belongs to, in-transit arrival dates |
| 14 | `14_wsa_align` | 1283 | Builds WSA_Daily (one layout per sheet, front / back glass split), vendor pools (ALLOWED VENDORS), deletes old outputs | `align_wsa`, `vendor_pools`, `vp_dead`, `ata_days`, `eta_days`, `delete_old_outputs` | WSA_Daily layout or formulas, allowed vendors, which old outputs get deleted |
| 15 | `15_main_run` | 674 | The main run: order of steps, console messages, saves every output file | (no functions: the `if __name__ == '__main__'` block) | run order, output file names, console / warning text, new output file |

## Where the rules are
- Short rules: `CLAUDE.md` (numbers register enforced by `python check_rules.py`).
- Full detail per feature: `docs/RULES_DETAIL.md` (read only the section you need).

## Rules for changes (same as before the split)
- A new rule goes into BOTH WSA_Daily (+ Inbound) AND MRP_Daily in the same change.
- A new rule often touches 2 files, for example 11 (the setting) + 07 (where it is used).
- After a change: `python check_rules.py` must PASS, then run before/after on the real inputs and compare every sheet.

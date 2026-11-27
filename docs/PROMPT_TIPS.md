# Prompt Tips (fewer messages, better code)

Goal: one clear message -> a tested result. Most extra messages come from missing facts, not from bad wording.

## 1. The 5 parts of a good prompt
1. **Symptom / goal**: what you see, on which sheet, for which material.
2. **Expected**: the value or behavior you want.
3. **Scope**: "fix the RULE for ALL materials", "both 2_WSA and 1_MRP".
4. **Constraints**: what must NOT change (numbers register, planner's WSA file, other sheets).
5. **Proof**: what you want to see back (before/after table, check_rules PASS, check_align PASS).

## 2. Bug-fix template (from CLAUDE.md, copy and fill)
> Symptom (example only): material ___, sheet ___, shows ___, expected ___. Find the root cause in the logic and fix the rule for
> ALL materials with the same pattern. No hardcoded codes. Apply to both 2_WSA and 1_MRP. Show a before/after table of every
> material that changed, plus any similar material that did not change and why. Build the recommended option.

Why it works: "example only" stops a one-material patch. "Recommended option" saves a question round.

## 3. New-feature template
> Goal: ___ (what the planner decides with it). Where: sheet/output number ___ (e.g. "fix 3 ...", '1_MRP 03'). Inputs used: ___.
> Rule: ___ (formula in words, with one worked example from my data). Do not change: ___. Show: ___.

## 4. Question template (no code change)
> Explain how ___ is calculated for material ___. Use the code, quote the function name, give a numeric example. Do not edit files.

## 5. Habits that save messages
- **Give one concrete example with real numbers**: "RMxxx, stock 1,200, plan 100/day -> expect 12 days, sheet shows 9".
- **Name the output by number**: "5_Web", "1_MRP 03", "4_Status". No guessing.
- **Say "build the recommended option"** when choices are open. Ask only when the answer changes numbers.
- **One topic per session**. Start a new session for a new feature or a new day's run (context is cleaner, usage is lower).
- **Batch related fixes** in one message as a numbered list; each item gets its own symptom + expected.
- **Attach / name the input files** the run needs (the 5 inputs) so the test uses real data.
- **State the rule once, in business words**: "stock never below 0", "lead time = threshold". Claude maps it to code.
- **Say what "done" means**: "done = run passes, before/after shown, dist file rebuilt".
- **Correct me**: if a result looks wrong, give the cell and the value you expected, not "it is wrong".
- **Ask for the plan first only for big changes**: "Plan only, no code, 5 bullets", then "go".

## 6. Bad vs good
| Bad | Good |
|---|---|
| "Fix the glass problem" | "Material 268 (back glass) on 2_WSA 'Glass ...' shows 14 days; expected 9 because stock 900 / usage 100. Fix the rule for all glass." |
| "It does not match" | "1_MRP ALERT row for RMxxx says 7 days, 4_Status says 9. check_align should PASS. Find why." |
| "Add a column" | "Add 'Supplier lead time' to 4_Status BOM Weekly after column H, from Settings TRANSIT LEAD TIME. Same value in 5_Web." |
| "Make it faster" | "Run takes ~70 s. Add a Timing line per step, show the 5 slowest, then fix the biggest one without changing outputs." |

## 7. Production-ready code checklist (ask for it)
- Edge cases: empty plan, zero stock, missing lead time, unknown unit, substitute material.
- No hardcoded material codes; settings come from the Settings sheet.
- Same logic in every output (one function per calculation).
- Test on real inputs, compare every sheet before/after, only expected cells change.
- Delete unused code. No new files unless needed.
- Python 3.9+ compatible.

## 8. Use the repo rules, don't repeat them
CLAUDE.md is loaded every session. Don't re-explain: top rules, numbers register, output layout. Say "per CLAUDE.md" and only add what is new.
When you decide a new rule, say "add this as a rule": one line in CLAUDE.md + detail in docs/RULES_DETAIL.md + check_rules.py.

## 9. Usage-limit tips
- Ask for the **short answer** ("one line on what changed") when you do not need the explanation.
- Avoid pasting big tables; give the 3 materials that show the problem.
- Do not ask Claude to re-read big files: name the function or sheet ("see src/07 ALERT").
- Long chat = costly chat. New session after a task is done.

---
name: xlsx
description: "Use this skill any time a spreadsheet file is the primary input or output. This means any task where the user wants to: open, read, edit, or fix an existing .xlsx, .xlsm, .csv, or .tsv file (e.g., adding columns, computing formulas, formatting, charting, cleaning messy data); create a new spreadsheet from scratch or from other data sources; or convert between tabular file formats. Trigger especially when the user references a spreadsheet file by name or path — even casually (like \"the xlsx in my downloads\") — and wants something done to it or produced from it. Also trigger for cleaning or restructuring messy tabular data files (malformed rows, misplaced headers, junk data) into proper spreadsheets. The deliverable must be a spreadsheet file. Do NOT trigger when the primary deliverable is a Word document, HTML report, standalone Python script, database pipeline, or Google Sheets API integration, even if tabular data is involved."
license: MIT
---

# Requirements for Outputs

## All Excel files

### Professional Font
- Use a consistent, professional font (e.g., Arial, Times New Roman) for all deliverables unless otherwise instructed by the user.

### Zero Formula Errors
- Every Excel file MUST be delivered with ZERO formula errors (`#REF!`, `#DIV/0!`, `#VALUE!`, `#N/A`, `#NAME?`, `#NUM!`, `#NULL!`, `#SPILL!`, `#CALC!`).
- `python3 <skill-dir>/scripts/xlsx.py validate <file>` MUST exit 0 before you report the task done.

### Preserve Existing Templates (when updating templates)
- Study and EXACTLY match existing format, style, and conventions when modifying files.
- Never impose standardized formatting on files with established patterns.
- Existing template conventions ALWAYS override these guidelines.

## Financial models

Unless otherwise stated by the user or by an existing template:

### Color Coding Standards (industry convention)
- **Blue text (RGB 0,0,255)** — hardcoded inputs, numbers users will change for scenarios.
- **Black text (RGB 0,0,0)** — all formulas and calculations.
- **Green text (RGB 0,128,0)** — links pulling from other worksheets within the same workbook.
- **Red text (RGB 255,0,0)** — external links to other files.
- **Yellow background (RGB 255,255,0)** — key assumptions needing attention.

### Number Formatting Standards
- **Years**: format as text strings ("2024" not "2,024").
- **Currency**: `$#,##0` format. ALWAYS specify units in headers ("Revenue ($mm)").
- **Zeros**: format so all zeros render as `-` (e.g. `"$#,##0;($#,##0);-"`).
- **Percentages**: default to `0.0%` (one decimal).
- **Multiples**: `0.0x` for valuation multiples (EV/EBITDA, P/E).
- **Negative numbers**: parentheses `(123)` not minus `-123`.

### Formula Construction Rules
- Put ALL assumptions (growth rates, margins, multiples, etc.) in dedicated assumption cells.
- Reference assumption cells in formulas — do not hardcode the value into the formula.
  - Use `=B5*(1+$B$6)` not `=B5*1.05`.
- Verify all cell references are correct; check off-by-one in ranges.
- Keep formulas consistent across projection periods.
- Test with edge cases (zero values, negative numbers).
- Verify no unintended circular references.

### Documenting Hardcoded Values
Add a comment to the cell, or put a note in the column next to the cell. Format:
```
Source: [System/Document], [Date], [Specific Reference], [URL if applicable]
```
Examples:
- `Source: Company 10-K, FY2024, Page 45, Revenue Note, [SEC EDGAR URL]`
- `Source: Bloomberg Terminal, 8/15/2025, AAPL US Equity`

# XLSX creation, editing, and analysis

## Overview

This skill ships an openpyxl-based toolchain with in-process formula recalc via [formualizer](https://pypi.org/project/formualizer/) (a Rust crate with Python bindings — no LibreOffice / Node / subprocess required). The four-step workflow below is mandatory; everything else is convenience.

Tools available in this skill:
- **openpyxl** — read/write/edit cells, formulas, formatting.
- **formualizer** — in-process formula evaluator (drives `recalc` / `validate`).
- **xlsx_kit** — save wrapper that produces xlsx + Univer JSON + optional snapshot event in one call.
- **xlsx.py** — unified CLI dispatching to inspect / audit / recalc / validate / extend-formula / doctor.

## Important Requirements

- Every edit happens through Python (`openpyxl`) running in a **bash** subprocess. NEVER use the agent's `write` or `edit` tools on `.xlsx` files — they treat the file as text and will corrupt the binary zip container.
- **Invoke python via `~/.local/share/pi-sheets/.venv/bin/python`** (created by `install.sh`) instead of bare `python3`. The venv has `openpyxl` + `formualizer` installed; bare `python3` may not. `xlsx.py` itself auto-redirects to this interpreter, but heredocs that the agent writes need to use the absolute path explicitly.
- Use `xlsx_kit.save(wb, path)` instead of `wb.save(path)`. `xlsx_kit.save` writes the xlsx, generates a Univer JSON sidecar if the bridge is available, and emits a `workbook_snapshot` event (no-op if no consumer is wired up).
- After ANY formula-bearing edit: `<skill-dir>/scripts/xlsx.py validate <file>`. Do NOT call `audit` and `recalc` separately — `validate` is the union and runs in one pass.
- If any script reports `ModuleNotFoundError`, run `<skill-dir>/scripts/xlsx.py doctor --install` once. It creates the venv at `~/.local/share/pi-sheets/.venv` and installs deps. Idempotent.

## CRITICAL: Use Formulas, Not Hardcoded Values

ALWAYS use Excel formulas instead of computing values in Python and writing the result. A static value is a bug — the spreadsheet must recalculate when source data changes.

### ❌ WRONG — Hardcoding Calculated Values

```python
# Bad: calculating in Python and hardcoding the result
total = sum(row['Sales'] for row in rows)
sheet['B10'] = total                # hardcodes 5000

# Bad: computing growth rate in Python
growth = (latest - earliest) / earliest
sheet['C5'] = growth                # hardcodes 0.15

# Bad: Python calculation for average
sheet['D20'] = sum(values) / len(values)   # hardcodes 42.5
```

### ✅ CORRECT — Using Excel Formulas

```python
sheet['B10'] = '=SUM(B2:B9)'
sheet['C5']  = '=(C4-C2)/C2'
sheet['D20'] = '=AVERAGE(D2:D19)'
```

This applies to ALL calculations — totals, percentages, ratios, differences, multiples. If you find yourself doing arithmetic in Python and assigning the result to a cell, stop and write a formula instead.

## Mandatory Workflow

1. **Orient first** — before any edit:
   ```bash
   python3 <skill-dir>/scripts/xlsx.py inspect <file>
   ```
   Read the JSON: sheet names, dimensions, formula counts. Decide which sheet(s) you need.

2. **Edit via openpyxl in a bash heredoc**:
   ```bash
   ~/.local/share/pi-sheets/.venv/bin/python <<'PY'
   import sys
   sys.path.insert(0, '<skill-dir>/scripts')
   import xlsx_kit
   from openpyxl import load_workbook

   wb = load_workbook('workbook.xlsx')
   ws = wb['Sheet1']
   ws['A1'] = 'Revenue'
   ws['B2'] = '=SUM(B3:B10)'
   xlsx_kit.save(wb, 'workbook.xlsx')   # NOT wb.save(...)
   PY
   ```
   Preserve every formula and named range the task does not explicitly remove.

3. **Save via `xlsx_kit.save`** — never `wb.save`. The wrapper produces the Univer JSON sidecar and emits the snapshot event in one call.

4. **Validate before finishing**:
   ```bash
   python3 <skill-dir>/scripts/xlsx.py validate <file>
   ```
   Exit 0 + `"ok": true` in the JSON is the only acceptable result. If errors appear, fix them and re-validate in the same workflow.

5. **Preserve structure** — do not delete sheets, rename columns, or remove formulas unless the task explicitly requires it.

## Creating new Excel files

```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
import sys
sys.path.insert(0, '<skill-dir>/scripts')
import xlsx_kit

wb = Workbook()
ws = wb.active
ws.title = 'Budget'

ws['A1'] = 'Category'; ws['B1'] = 'Planned'; ws['C1'] = 'Actual'; ws['D1'] = 'Variance'
ws.append(['Rent', 1000, 1100, '=C2-B2'])
ws.append(['Food',  300,  250, '=C3-B3'])
ws['A4'] = 'Total'
ws['B4'] = '=SUM(B2:B3)'
ws['C4'] = '=SUM(C2:C3)'
ws['D4'] = '=SUM(D2:D3)'

ws['A1'].font = Font(bold=True)
ws['A1'].fill = PatternFill('solid', start_color='FFFF00')
ws.column_dimensions['A'].width = 20

xlsx_kit.save(wb, 'output.xlsx')
```

## Editing existing Excel files

```python
from openpyxl import load_workbook
import sys
sys.path.insert(0, '<skill-dir>/scripts')
import xlsx_kit

wb = load_workbook('existing.xlsx')        # keeps formulas as strings
ws = wb.active                              # or wb['SheetName']

for name in wb.sheetnames:
    print(name, wb[name].max_row, wb[name].max_column)

ws['A1'] = 'New header'
ws.insert_rows(2)
ws.delete_cols(3)

new = wb.create_sheet('Notes')
new['A1'] = 'Source data'

xlsx_kit.save(wb, 'existing.xlsx')
```

## Recalculating formulas

openpyxl writes formula text but NOT computed values. Use `xlsx.py validate` to recalc and audit in one pass:

```bash
python3 <skill-dir>/scripts/xlsx.py validate <file>
```

JSON output shape:

```json
{
  "ok": true,
  "path": "/abs/path/wb.xlsx",
  "audit": {
    "formula_cells": 42,
    "formula_error_count": 0,
    "ref_error_count": 0,
    "formula_errors": [],
    "ref_errors": []
  },
  "recalc": {
    "formula_cells": 42,
    "formula_error_count": 0,
    "errors": []
  },
  "elapsed_ms": 18
}
```

If `ok: false`, `audit.formula_errors` and `recalc.errors` give you `{sheet, cell, formula, ...}` records pointing at exact failure cells. Fix and re-run `validate`.

Common errors and fixes:
- `#REF!` — formula references a deleted/missing cell. Fix the reference.
- `#DIV/0!` — denominator evaluates to zero. Guard with `IFERROR` or check inputs.
- `#VALUE!` — wrong type passed to a function (text where number expected, etc.).
- `#NAME?` — typo in function name, or function unsupported by formualizer.
- `#N/A` — `VLOOKUP`/`MATCH` couldn't find a value.

## Formula Verification Checklist

Before declaring success:

### Essential Verification
- [ ] **Test 2-3 sample references** before building the full model.
- [ ] **Column letter mapping** is right (column 64 = `BL`, not `BK`).
- [ ] **Row offset** is right (openpyxl rows are 1-indexed; if you computed from a 0-indexed source, add 1).
- [ ] **Validate exits 0** with zero errors.

### Common Pitfalls
- [ ] `None` / null values: check with `if value is not None` not truthy tests.
- [ ] Far-right columns: FY data often sits in columns 50+.
- [ ] Multiple matches: search all occurrences, not just the first.
- [ ] Division: check denominators are non-zero before writing `/` in formulas.
- [ ] Cross-sheet references: use `'Sheet With Spaces'!A1` quoting when the sheet name has spaces.

### Testing Strategy
- [ ] Start small — write 2–3 formula cells, validate, then expand.
- [ ] Verify all referenced cells actually exist.
- [ ] Test edge cases (zero, negative, very large values).

## Best Practices

### Working with openpyxl
- Cell indices are 1-based: `ws.cell(row=1, column=1)` is `A1`.
- Use `data_only=True` to **read** computed values: `load_workbook('file.xlsx', data_only=True)`.
- **WARNING:** if you open with `data_only=True` and save, formulas are replaced with values and **permanently lost**. Only open with `data_only=True` for read-only inspection.
- For large files: `read_only=True` for reading, `write_only=True` for streaming writes.
- Formulas are preserved across open/save cycles (without `data_only`) but NOT evaluated — `xlsx.py validate` is the recalc step.

### Working with xlsx_kit
- Always `xlsx_kit.save(wb, path)` — not `wb.save(path)`. The wrapper is the integration point with the Univer viewer and the event wire.
- `xlsx_kit.save` is safe to call when no consumer is configured — it falls back to a plain save plus a no-op event emit.

### Working with formualizer (via xlsx.py)
- `recalc` and `validate` use formualizer's in-process evaluator. No LibreOffice / Node / network. Sub-second for files with thousands of formulas.
- Functions formualizer supports: most of the common Excel surface (`SUM`, `IF`, `VLOOKUP`, `INDEX`, `MATCH`, `SUMIF`, `AVERAGEIF`, dates, text, financial). If you hit `#NAME?` on a function name, replace with an equivalent using supported primitives.

## Pure-function imports (composition)

For multi-step flows where you want branching/looping in one bash call, import the operation modules directly instead of invoking the CLI in subprocesses:

```python
import sys; sys.path.insert(0, '<skill-dir>/scripts')
from inspect_xlsx import inspect
from validate_xlsx import validate
from extend_formula import extend_formula
import xlsx_kit
from openpyxl import load_workbook

info = inspect('wb.xlsx')
if info['sheet_count'] > 1:
    extend_formula('wb.xlsx', source_cell='J16', target_range='K16:O16', overwrite=True)
    report = validate('wb.xlsx')
    assert report['ok'], report
```

Each module also exports its typed error class (`InspectError`, `AuditError`, `RecalcError`, `ValidateError`, `ExtendFormulaError`) with a `.code` attribute for catching specific failures.

## Operations reference

Every operation exists in three forms: a unified-CLI subcommand, a per-script CLI, and an importable Python function. Pick by use:
- **Unified CLI** (`xlsx.py <sub>`) — one-shot bash invocations.
- **Per-script CLI** (`inspect_xlsx.py`, etc.) — backward-compat; identical behavior.
- **Python function** — composition inside a heredoc.

### `inspect` — workbook structure summary

```python
from inspect_xlsx import inspect, InspectError
result = inspect(wb_path: str | Path) -> dict
```

CLI: `python3 <skill-dir>/scripts/xlsx.py inspect <wb>`

Returns:
```json
{
  "ok": true,
  "path": "/abs/path/wb.xlsx",
  "sheet_count": 2,
  "sheets": [
    {
      "name": "Budget",
      "dims": "A1:D6",
      "rows": 6,
      "cols": 4,
      "formula_cells": 7,
      "data_cells": 17,
      "first_row": 1,
      "first_row_values": ["Category", "Planned", "Actual", "Variance"]
    }
  ],
  "summary": "Budget(A1:D6,7f)"
}
```

Raises `InspectError(code=...)` with `code` ∈ `{not_found, bad_extension, open_failed}`.

### `audit` — formula audit (broken refs + cached errors)

```python
from audit_xlsx import audit, AuditError
result = audit(wb_path: str | Path) -> dict
```

CLI: `python3 <skill-dir>/scripts/xlsx.py audit <wb>`

Returns:
```json
{
  "ok": true,
  "path": "...",
  "sheet_count": 1,
  "formula_cells": 42,
  "formula_error_count": 0,
  "ref_error_count": 0,
  "sheets": [{"sheet": "...", "max_row": 6, "max_column": 4, "formula_cells": 7, "formula_error_cells": 0, "ref_error_formulas": 0}],
  "formula_errors": [{"sheet": "...", "cell": "B5", "formula": "=...", "cached_value": "#DIV/0!"}],
  "ref_errors":     [{"sheet": "...", "cell": "B5", "formula": "=#REF!"}]
}
```

Audits the cached `data_only` values for excel error sentinels and scans formula text for `#REF!`. Does NOT re-evaluate — use `recalc`/`validate` for that.

Raises `AuditError(code=...)` with `code` ∈ `{not_found, bad_extension, open_failed}`.

### `recalc` — in-process formula evaluation via formualizer

```python
from recalc_xlsx import recalc, RecalcError
result = recalc(
    wb_path: str | Path,
    write_evaluated: Path | None = None,   # optional: write value-only sidecar
) -> dict
```

CLI: `python3 <skill-dir>/scripts/xlsx.py recalc <wb> [--write-evaluated <out>]`

Returns:
```json
{
  "ok": true,
  "path": "...",
  "formula_cells": 42,
  "formula_error_count": 0,
  "errors": [{"sheet": "...", "coord": "B5", "formula": "=...", "value": "#DIV/0!"}],
  "elapsed_ms": 17,
  "evaluated_path": "/abs/out.xlsx",        // only when --write-evaluated set
  "evaluated_path_error": null              // only on failure to write sidecar
}
```

Raises `RecalcError(code=...)` with `code` ∈ `{not_found, bad_extension, open_failed, eval_failed}`.

### `validate` — audit + recalc combined (PREFER THIS)

```python
from validate_xlsx import validate, ValidateError
report = validate(wb_path: str | Path) -> dict
```

CLI: `python3 <skill-dir>/scripts/xlsx.py validate <wb>`

Returns the combined report — see "Recalculating formulas" section above for full shape.

Raises `ValidateError(code=...)` with `code` ∈ `{not_found, bad_extension, open_failed}`, or re-raises `RecalcError` from the recalc step.

### `extend_formula` — fill-right / fill-down with A1 shifts

```python
from extend_formula import extend_formula, ExtendFormulaError
result = extend_formula(
    wb_path: str | Path,
    source_cell: str,                     # e.g. "J16"
    target_range: str,                    # e.g. "K16:O16"
    sheet: str | None = None,             # defaults to active sheet
    overwrite: bool = False,              # default: skip non-empty cells
) -> dict
```

CLI: `python3 <skill-dir>/scripts/xlsx.py extend-formula <wb> --source-cell <A1> --target-range <A1:B5> [--sheet <name>] [--overwrite]`

Wraps `openpyxl.formula.translate.Translator` — copies the formula from `source_cell` across `target_range`, shifting A1 references as Excel does on fill operations. Saves the workbook in place when any cell is written.

Returns:
```json
{
  "ok": true,
  "sheet": "Sales",
  "source_cell": "J16",
  "source_formula": "=SUM(J2:J15)",
  "written": [{"coord": "K16", "formula": "=SUM(K2:K15)"}, ...],
  "skipped": [{"coord": "J16", "reason": "is source cell"}, ...],
  "errors":  [{"coord": "...", "error": "..."}]
}
```

Raises `ExtendFormulaError(code=...)` with `code` ∈ `{not_found, bad_extension, open_failed, sheet_not_found, source_not_formula, bad_range, save_failed}`.

### `xlsx_kit.save` — canonical save wrapper

```python
import xlsx_kit
xlsx_kit.save(wb: openpyxl.Workbook, path: str) -> None
```

No CLI form — invoked from Python only. Always use this instead of `wb.save(path)`.

Side effects:
- Writes the xlsx (always).
- Generates `<basename>.univer.json` sidecar when `UNIVER_BRIDGE_DIR` env points at the Univer conversion bridge (skipped silently otherwise).
- Emits a `workbook_snapshot` event (and a `workbook_snapshot_warning` if the bridge is unavailable) to `XLSX_EVENTS_PATH` (no-op if unset).

Safe to call in all host environments — no required env vars.

### `doctor` — python env preflight

```python
# CLI only:
python3 <skill-dir>/scripts/xlsx.py doctor
```

No importable form. Exits 0 with `status: ok` when both `openpyxl` and `formualizer` import cleanly; exits 1 otherwise with the missing dep and the exact `pip install -r requirements.txt` command. Always prints `sys.executable` so you can see which python the host picked up.

### `event_emit` — optional structured-event sink (advanced)

```python
from event_emit import event, emit_typed
event(kind: str, **payload)               # → {type: "script_event", kind, payload, ts, parentCallId}
emit_typed(type_: str, **fields)          # → {type: type_, ts, parentCallId, **fields}
```

Used internally by `xlsx_kit.save`. Both functions are no-ops when `XLSX_EVENTS_PATH` env is unset, so calling them from your own python is safe. See `scripts/EVENTS.md` for the wire contract.

## CLI reference

```bash
python3 <skill-dir>/scripts/xlsx.py <subcommand> [args]
```

| Subcommand | Purpose |
|---|---|
| `inspect <wb>` | Sheets + dims + formula counts (JSON) |
| `audit <wb>` | Formula audit only |
| `recalc <wb> [--write-evaluated <out>]` | Recalc via formualizer; `--write-evaluated` writes a value-only sidecar |
| `validate <wb>` | Audit + recalc combined — **always prefer this** |
| `extend-formula <wb> --source-cell <A1> --target-range <A1:B5> [--sheet <name>] [--overwrite]` | Fill-right / fill-down with A1 shifts |
| `doctor` | Preflight: verify python env + deps; prints which python the host is using |

Each subcommand also exists as a standalone per-script CLI (`inspect_xlsx.py`, etc.) for backwards compatibility. Prefer `xlsx.py <subcommand>`.

## Code Style Guidelines

**For Python code you generate:**
- Write minimal, concise code. No unnecessary comments.
- Avoid verbose variable names and redundant operations.
- Avoid unnecessary `print()` statements — the validate step is the source of truth, not print output.

**For Excel files themselves:**
- Add cell comments for complex formulas and important assumptions.
- Document data sources for hardcoded values (see "Documenting Hardcoded Values" above).
- Include section notes for key calculations.

## Caveats from real runs

Non-obvious failure modes observed in agentic eval runs. Follow them.

### NEVER use the `write` or `edit` tools on `.xlsx` files
The host's text-editing tools treat files as UTF-8. An `.xlsx` is a zip archive of XML. Writing markdown / text via `write` produces a 326-byte non-zip file that `openpyxl` cannot open — task fails with `BadZipFile: File is not a zip file`. **Only edit through `bash` → `python3` → `openpyxl`.**

### Always use a bash heredoc for multi-step edits
Single-line `python3 -c '...'` is fine for trivial reads, but multi-step edit + save + validate flows are dramatically more reliable as one heredoc:
```bash
~/.local/share/pi-sheets/.venv/bin/python <<'PY'
... full flow ...
PY
```
One LLM turn, one bash call, no fragile string escaping.

### `xlsx_kit.save`, not `wb.save`
`wb.save` works but skips Univer JSON conversion and the snapshot event. `xlsx_kit.save` is the canonical save call in this skill. The wrapper is safe even when no consumer is wired up — it falls back gracefully.

### `validate` is `audit + recalc` combined
Don't call `audit_xlsx.py` and `recalc_xlsx.py` separately. `validate` does both in one workbook walk and is materially faster. Calling both individually is wasted work and yields the same answer.

### `ModuleNotFoundError`? Run `doctor`
```bash
python3 <skill-dir>/scripts/xlsx.py doctor
```
It prints `sys.executable` and the missing dep — copy/paste the suggested `pip install` and retry. Most "missing dep" failures are actually the host launching from a shell whose `PATH` doesn't point at the venv with the deps installed; `doctor` makes that obvious.

### `<skill-dir>` is the absolute path of this `SKILL.md`'s parent directory
Resolve it concretely (e.g. `/Users/x/.claude/skills/xlsx`). Do NOT use a partial path. Sub-models that hallucinate `xlsx_agent/scripts` or `~/scripts` will get `ModuleNotFoundError`.

### Optional event wire
The skill emits structured events (`workbook_snapshot`, `script_event`) when `XLSX_EVENTS_PATH=<file>` is exported. Off by default — set it before launching the host to enable. See `scripts/EVENTS.md` for the schema.

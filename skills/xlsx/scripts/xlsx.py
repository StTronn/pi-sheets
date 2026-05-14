#!/usr/bin/env python3
"""xlsx — unified CLI for xlsx skill operations.

Dispatches to the same pure functions the per-script CLIs expose. JSON to
stdout; exit 0 on success, 1 on any error.

Usage:
    python3 xlsx.py inspect <wb.xlsx>
    python3 xlsx.py audit <wb.xlsx>
    python3 xlsx.py recalc <wb.xlsx> [--write-evaluated <out.xlsx>]
    python3 xlsx.py validate <wb.xlsx>
    python3 xlsx.py extend-formula <wb.xlsx> --source-cell <A1>
                    --target-range <A1:B5> [--sheet <name>] [--overwrite]
    python3 xlsx.py doctor

Library use: prefer importing the operation modules directly
    (`from inspect_xlsx import inspect`, etc.) over invoking this CLI in a
    subprocess from python.

Note: operation modules are imported lazily inside each subcommand handler so
that `doctor` works even when openpyxl / formualizer are missing. The doctor
is the right thing to run when a `ModuleNotFoundError` appears elsewhere.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

# Self-bootstrap: ensure sibling modules import without host PYTHONPATH setup.
sys.path.insert(0, str(Path(__file__).resolve().parent))


def _print(result: dict) -> int:
    print(json.dumps(result))
    return 0 if result.get("ok", True) else 1


def _fail(msg: str, *, code: str = "error") -> int:
    print(json.dumps({"ok": False, "code": code, "message": msg}))
    return 1


def _missing_dep_hint(exc: ModuleNotFoundError) -> int:
    skill_dir = Path(__file__).resolve().parent.parent
    return _fail(
        f"missing python dependency: {exc.name}. "
        f"Run `python3 {Path(__file__).resolve()} doctor` for details, "
        f"or install: pip install -r {skill_dir}/requirements.txt",
        code="missing_dependency",
    )


def _cmd_inspect(args: argparse.Namespace) -> int:
    try:
        from inspect_xlsx import inspect, InspectError
    except ModuleNotFoundError as exc:
        return _missing_dep_hint(exc)
    try:
        return _print(inspect(args.path))
    except InspectError as exc:
        return _fail(str(exc), code=exc.code)


def _cmd_audit(args: argparse.Namespace) -> int:
    try:
        from audit_xlsx import audit, AuditError
    except ModuleNotFoundError as exc:
        return _missing_dep_hint(exc)
    try:
        return _print(audit(args.path))
    except AuditError as exc:
        return _fail(str(exc), code=exc.code)


def _cmd_recalc(args: argparse.Namespace) -> int:
    try:
        from recalc_xlsx import recalc, RecalcError
    except ModuleNotFoundError as exc:
        return _missing_dep_hint(exc)
    try:
        result = recalc(
            args.path,
            write_evaluated=Path(args.write_evaluated) if args.write_evaluated else None,
        )
        return _print(result)
    except RecalcError as exc:
        return _fail(str(exc), code=exc.code)


def _cmd_validate(args: argparse.Namespace) -> int:
    try:
        from validate_xlsx import validate, ValidateError
        from recalc_xlsx import RecalcError
    except ModuleNotFoundError as exc:
        return _missing_dep_hint(exc)
    try:
        return _print(validate(args.path))
    except (ValidateError, RecalcError) as exc:
        return _fail(str(exc), code=exc.code)


def _cmd_extend_formula(args: argparse.Namespace) -> int:
    try:
        from extend_formula import extend_formula, ExtendFormulaError
    except ModuleNotFoundError as exc:
        return _missing_dep_hint(exc)
    try:
        result = extend_formula(
            args.path,
            source_cell=args.source_cell,
            target_range=args.target_range,
            sheet=args.sheet,
            overwrite=args.overwrite,
        )
        return _print(result)
    except ExtendFormulaError as exc:
        return _fail(str(exc), code=exc.code)


def _cmd_doctor(_args: argparse.Namespace) -> int:
    # `doctor` is the one subcommand that must work even when deps are missing,
    # since its whole job is to diagnose missing deps.
    import doctor as doctor_mod
    return doctor_mod.main()


def main() -> int:
    p = argparse.ArgumentParser(prog="xlsx", description="Unified xlsx skill CLI.")
    sub = p.add_subparsers(dest="cmd", required=True)

    pi = sub.add_parser("inspect", help="Workbook structure summary.")
    pi.add_argument("path"); pi.set_defaults(fn=_cmd_inspect)

    pa = sub.add_parser("audit", help="Audit formulas for broken refs / cached errors.")
    pa.add_argument("path"); pa.set_defaults(fn=_cmd_audit)

    pr = sub.add_parser("recalc", help="Recalculate formulas via formualizer.")
    pr.add_argument("path"); pr.add_argument("--write-evaluated", metavar="PATH")
    pr.set_defaults(fn=_cmd_recalc)

    pv = sub.add_parser("validate", help="Audit + recalc in one pass (preferred).")
    pv.add_argument("path"); pv.set_defaults(fn=_cmd_validate)

    pe = sub.add_parser("extend-formula", help="Copy a formula across a target range (A1 shifts).")
    pe.add_argument("path"); pe.add_argument("--source-cell", required=True)
    pe.add_argument("--target-range", required=True); pe.add_argument("--sheet")
    pe.add_argument("--overwrite", action="store_true")
    pe.set_defaults(fn=_cmd_extend_formula)

    pd = sub.add_parser("doctor", help="Preflight: verify python env + deps.")
    pd.set_defaults(fn=_cmd_doctor)

    args = p.parse_args()
    return args.fn(args)


if __name__ == "__main__":
    sys.exit(main())

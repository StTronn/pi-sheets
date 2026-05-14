#!/usr/bin/env python3
"""Verify the xlsx skill's python runtime is wired up.

Exits 0 if every required dep imports cleanly. Exits 1 with an actionable hint
otherwise, printing the `sys.executable` so users with multiple pythons can
see which one the host actually picked up.

With `--install`, shells out to install.sh to create a venv + install deps
automatically. Idempotent — safe to re-run.

Run me first when scripts fail with ModuleNotFoundError.
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path

REQUIRED = [
    ('openpyxl', 'openpyxl>=3.1'),
    ('formualizer', 'formualizer>=0.5.8'),
]


def _check_deps() -> list[str]:
    missing: list[str] = []
    for mod, spec in REQUIRED:
        try:
            __import__(mod)
        except ImportError:
            missing.append(spec)
    return missing


def _report(*, install: bool) -> int:
    skill_dir = Path(__file__).resolve().parent.parent
    install_sh = skill_dir / 'scripts' / 'install.sh'

    missing = _check_deps()
    print(f"python: {sys.executable}")
    print(f"skill:  {skill_dir}")

    if not missing:
        print("status: ok")
        return 0

    print(f"missing: {', '.join(missing)}")

    if install:
        # Run install.sh — it picks a working python, creates a venv at
        # ~/.local/share/pi-sheets/.venv, installs deps, and caches the path.
        print(f"running: bash {install_sh}")
        rc = subprocess.call(['bash', str(install_sh)])
        return rc

    print(f"fix:     bash {install_sh}")
    print(f"or:      pip install -r {skill_dir}/requirements.txt")
    return 1


def main() -> int:
    p = argparse.ArgumentParser(prog='doctor')
    p.add_argument(
        '--install', action='store_true',
        help='Auto-install deps via install.sh (creates a venv if needed).',
    )
    args = p.parse_args()
    return _report(install=args.install)


if __name__ == '__main__':
    sys.exit(main())

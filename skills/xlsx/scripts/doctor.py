#!/usr/bin/env python3
"""Verify the xlsx skill's python runtime is wired up.

Exits 0 if every required dep imports cleanly. Exits 1 with an actionable hint
otherwise, printing the `sys.executable` so users with multiple pythons can
see which one the host actually picked up.

Run me first when scripts fail with ModuleNotFoundError.
"""

from __future__ import annotations

import os
import sys

REQUIRED = [
    ('openpyxl', 'openpyxl>=3.1'),
    ('formualizer', 'formualizer>=0.5.8'),
]


def main() -> int:
    missing: list[str] = []
    for mod, spec in REQUIRED:
        try:
            __import__(mod)
        except ImportError:
            missing.append(spec)

    skill_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    print(f"python: {sys.executable}")
    print(f"skill:  {skill_dir}")

    if missing:
        print(f"missing: {', '.join(missing)}")
        print(f"fix:     pip install -r {skill_dir}/requirements.txt")
        return 1

    print("status: ok")
    return 0


if __name__ == '__main__':
    sys.exit(main())

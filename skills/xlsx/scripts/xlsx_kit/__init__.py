"""Thin save-side wrapper over openpyxl.

Currently exposes one operation: `save(wb, path)`. It writes the workbook,
runs Univer conversion in-process, and emits a `workbook_snapshot` event
via event_emit. Replaces the TS sha256-diff post-hook — the lib already knows
when something changed and can emit directly.

Future ops (`set`, `range`, `validate`, ...) will live alongside `save` and
follow the same auto-emit pattern.
"""

from __future__ import annotations

import hashlib
import json
import os
import sys
from typing import Any

# Self-bootstrap: ensure sibling `event_emit.py` (one level up in scripts/) is
# importable without the host setting PYTHONPATH.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from event_emit import event, emit_typed


def save(wb: Any, path: str) -> None:
    """Save workbook, convert to Univer JSON, emit workbook_snapshot."""
    pre = _hash(path) if os.path.exists(path) else ''
    wb.save(path)
    post = _hash(path)

    univer_path = _strip_xlsx(path) + '.univer.json'
    try:
        _convert_to_univer(path, univer_path)
    except Exception as e:
        event('workbook_snapshot_warning', xlsxPath=path, error=str(e))
        univer_path = ''

    emit_typed(
        'workbook_snapshot',
        xlsxPath=path,
        univerPath=univer_path,
        preHash=pre,
        postHash=post,
    )


def _hash(path: str) -> str:
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(64 * 1024), b''):
            h.update(chunk)
    return h.hexdigest()


def _strip_xlsx(path: str) -> str:
    return path[:-5] if path.lower().endswith('.xlsx') else path


def _convert_to_univer(xlsx_path: str, out_path: str) -> None:
    bridge = os.environ.get('UNIVER_BRIDGE_DIR')
    if not bridge:
        raise RuntimeError('UNIVER_BRIDGE_DIR env not set')
    if bridge not in sys.path:
        sys.path.insert(0, bridge)
    from excel_to_univer import convert  # type: ignore
    data = convert(xlsx_path)
    with open(out_path, 'w') as f:
        json.dump(data, f)

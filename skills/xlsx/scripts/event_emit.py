"""Harness-agnostic structured event emitter.

Producer writes one JSON line per event to ``XLSX_EVENTS_PATH`` (set by the
host once per session). Optionally tags each line with ``XLSX_PARENT_CALL_ID``
(set by the host per parent tool invocation when available). No-op if no event
path is set, so scripts remain safe to run standalone.

Two entry points:
    event(kind, **payload)       → {type:'script_event', kind, payload}
    emit_typed(type_, **fields)  → top-level typed event (e.g. workbook_snapshot)

Wire contract: see EVENTS.md.

Concurrency: a single-line write to an O_APPEND file is atomic for sizes
< PIPE_BUF (4 KiB on macOS+Linux). Consumers tailing the same file via
append-mode writes can safely interleave.
"""

from __future__ import annotations

import json
import os
import time
from typing import Any


def _path() -> str | None:
    return os.environ.get('XLSX_EVENTS_PATH')


def _envelope() -> dict[str, Any]:
    return {
        'ts': int(time.time() * 1000),
        'parentCallId': os.environ.get('XLSX_PARENT_CALL_ID'),
    }


def _write(record: dict[str, Any]) -> None:
    p = _path()
    if not p:
        return
    line = json.dumps(record, separators=(',', ':'), default=str)
    with open(p, 'a') as f:
        f.write(line + '\n')


def event(kind: str, **payload: Any) -> None:
    """Emit a generic kind-tagged script_event."""
    _write({'type': 'script_event', **_envelope(), 'kind': kind, 'payload': payload})


def emit_typed(type_: str, **fields: Any) -> None:
    """Emit a top-level typed event (e.g. 'workbook_snapshot')."""
    _write({'type': type_, **_envelope(), **fields})

# xlsx skill — event wire contract

Harness-agnostic JSONL stream emitted by skill scripts. Any host that wants to
observe workbook changes or script telemetry sets two env vars and tails a file.

## Activation

| Env var                 | Set by | Purpose                                                  |
|-------------------------|--------|----------------------------------------------------------|
| `XLSX_EVENTS_PATH`      | host   | Absolute path to wire file. Unset → emitter is a no-op.  |
| `XLSX_PARENT_CALL_ID`   | host   | Per-tool-call id stamped on every line for correlation.  |

Host creates the file (or `touch`es it) once per session and tails it. Producer
opens in append mode per write.

## Line format

One JSON object per line, no trailing whitespace:

```json
{"type":"<event-type>","ts":1715000000000,"parentCallId":"call_abc",...}
```

Fields always present:
- `type` — discriminator (see below)
- `ts` — emitter wall-clock, ms since epoch
- `parentCallId` — value of `XLSX_PARENT_CALL_ID` at emit time, or `null`

## Event types

### `script_event` — generic kind-tagged log
```json
{"type":"script_event","kind":"recalc.start","payload":{"sheet":"P&L"}}
```
Emitted by `event_emit.event(kind, **payload)`.

### `workbook_snapshot` — emitted by `xlsx_kit.save`
```json
{"type":"workbook_snapshot","xlsxPath":"/abs/wb.xlsx","univerPath":"/abs/wb.univer.json","changed":true,"shaBefore":"...","shaAfter":"..."}
```
Host UIs (e.g. Univer viewer) consume this to refresh without re-reading the
xlsx. `changed:false` lines are emitted on idempotent saves.

Additional typed events may be added via `event_emit.emit_typed(type_, **fields)`.

## Concurrency

POSIX `O_APPEND` writes < `PIPE_BUF` (4 KiB on macOS/Linux) are atomic. Keep
individual lines under 4 KiB; for larger payloads, write a side file and emit
a `script_event` pointing at it.

## Standalone runs

If `XLSX_EVENTS_PATH` is unset, all emit calls are no-ops. Scripts remain
runnable from a plain shell with no host attached.

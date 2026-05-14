# xlsx skill bundle

A self-contained spreadsheet-editing skill for agentic coding hosts (Pi,
Claude Code, Codex). Wraps `openpyxl` + the `formualizer` calc engine behind
a mandatory four-step workflow (inspect → edit → save via `xlsx_kit` →
validate), exposing every operation as both a CLI subcommand and an
importable python module.


## What's inside

```
skill/
├── SKILL.md              Agent-facing workflow (loaded by host)
├── README.md             This file (humans)
├── requirements.txt      Python deps: openpyxl, formualizer
└── scripts/
    ├── xlsx.py           Unified CLI (inspect / audit / recalc / validate / extend-formula / doctor)
    ├── inspect_xlsx.py   Pure function + per-script CLI: workbook structure summary
    ├── audit_xlsx.py     Formula audit (broken refs, cached errors)
    ├── recalc_xlsx.py    Recalc via formualizer
    ├── validate_xlsx.py  Audit + recalc combined — preferred
    ├── extend_formula.py
    ├── doctor.py         Preflight: python env + dep check
    ├── event_emit.py     Optional wire emitter (no-op unless XLSX_EVENTS_PATH set)
    ├── EVENTS.md         Wire contract for the optional event sink
    └── xlsx_kit/         save() wrapper (xlsx + univer.json + snapshot event)
```



## Install — Pi coding-agent

Use the parent `pi-sheets` package. The tiny extension in `../src/index.ts`
registers this directory via `resources_discover`:

```bash
# from npm
pi install npm:pi-sheets

# or directly from git
pi install git:github.com/StTronn/pi-sheets

pip install -r ~/.pi/agent/extensions/pi-sheets/skill/requirements.txt
```

See `../README.md` for full docs.

## Install — Claude Code

```bash
# clone the parent repo once
git clone https://github.com/StTronn/pi-sheets ~/code/pi-sheets

# symlink THIS DIRECTORY into Claude Code's skill discovery path
mkdir -p ~/.claude/skills
ln -s ~/code/pi-sheets/skill ~/.claude/skills/xlsx

# install python deps (any venv works; activate before launching `claude`)
pip install -r ~/code/pi-sheets/skill/requirements.txt
```

Claude Code auto-discovers the skill on next launch. The `description:` field
in `SKILL.md` is what triggers it — say "edit this xlsx" or similar and Claude
will load the workflow.

For project-local installs instead of global, symlink into
`<your-project>/.claude/skills/xlsx`.

## Install — Codex

```bash
git clone https://github.com/StTronn/pi-sheets ~/code/pi-sheets

mkdir -p ~/.codex/skills
ln -s ~/code/pi-sheets/skill ~/.codex/skills/xlsx

pip install -r ~/code/pi-sheets/skill/requirements.txt
```

Same auto-discovery model. Project-local equivalent:
`<your-project>/.codex/skills/xlsx`.

## Python setup

Python ≥3.10 with `openpyxl` and `formualizer` on the path. A project-local
venv works well:

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```

Launch the host (`claude` / `codex` / `pi`) from the shell with the venv
active. The scripts inherit `PATH` and pick up the right interpreter — no
PYTHONPATH tweaks needed; `xlsx_kit/__init__.py` self-bootstraps imports for
its sibling modules.

Stuck on `ModuleNotFoundError`? Run the doctor:

```bash
python3 scripts/xlsx.py doctor
```

It prints the exact python the host is using and the exact `pip install`
command to fix any missing deps.

## Optional: event wire

The skill emits structured events (`workbook_snapshot`, `script_event`) for
hosts that want to observe progress (e.g. drive a live workbook viewer).
**Off by default.** Enable per session:

```bash
export XLSX_EVENTS_PATH=/tmp/xlsx-events.ndjson
claude   # or: codex / pi
```

Tail the file in another terminal:

```bash
tail -F /tmp/xlsx-events.ndjson | jq .
```

Full wire schema: [`scripts/EVENTS.md`](scripts/EVENTS.md).

## Updating

Skill installs are symlinks — `git pull` inside `~/code/pi-sheets` updates
every host at once. No reinstall step.

## License

MIT — see `../LICENSE`.

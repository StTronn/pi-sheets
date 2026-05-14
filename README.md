# pi-sheets

A spreadsheet-editing skill for agentic coding hosts — orient → edit via
`openpyxl` → save via `xlsx_kit` → validate, preserving formulas and structure.
Works in **Pi coding-agent**, **Claude Code**, and **OpenAI Codex**.

Wraps two libraries:
- **[openpyxl](https://openpyxl.readthedocs.io/)** for workbook IO
- **[formualizer](https://pypi.org/project/formualizer/)** (Rust + Python bindings) for in-process formula recalc

## Install

### Pi coding-agent

```bash
# from npm (recommended)
pi install npm:@sttronn/pi-sheets

# or directly from git
pi install git:github.com/StTronn/pi-sheets

# find the requirements.txt path + install python deps in one step
pi exec -- python3 -c "
import sys, glob, os
for p in glob.glob(os.path.expanduser('~/.pi/agent') + '/**/pi-sheets/skills/xlsx/requirements.txt', recursive=True):
    print(p)
" | xargs pip install -r
```

If you'd rather see the resolved path yourself first, in a new pi session run
`/xlsx-doctor` or call the bundled doctor directly (it prints `sys.executable`
+ the missing dep + the exact `pip install` command):

```bash
python3 ~/.pi/agent/git/github.com/StTronn/pi-sheets/skills/xlsx/scripts/xlsx.py doctor
# or for npm installs:
python3 ~/.pi/agent/packages/@sttronn/pi-sheets/skills/xlsx/scripts/xlsx.py doctor
```

Pi runs the tiny extension in `src/index.ts`, which contributes `skills/` via
`resources_discover`. The agent automatically loads `SKILL.md` and follows the
workflow.

### Claude Code

```bash
git clone https://github.com/StTronn/pi-sheets ~/code/pi-sheets

mkdir -p ~/.claude/skills
ln -s ~/code/pi-sheets/skills/xlsx ~/.claude/skills/xlsx

pip install -r ~/code/pi-sheets/skills/xlsx/requirements.txt
```

For project-local: symlink to `<your-project>/.claude/skills/xlsx`.

### OpenAI Codex

```bash
git clone https://github.com/StTronn/pi-sheets ~/code/pi-sheets

mkdir -p ~/.codex/skills
ln -s ~/code/pi-sheets/skills/xlsx ~/.codex/skills/xlsx

pip install -r ~/code/pi-sheets/skills/xlsx/requirements.txt
```

For project-local: symlink to `<your-project>/.codex/skills/xlsx`.

### Manual / development install (any host)

```bash
git clone https://github.com/StTronn/pi-sheets ~/code/pi-sheets
cd ~/code/pi-sheets && npm install            # only needed for typechecking src/
pip install -r skills/xlsx/requirements.txt
```

Symlink `skills/xlsx/` into the host's skill discovery dir as shown above.

## Repo layout

```
pi-sheets/
├── skills/
│   └── xlsx/     ← host-agnostic skill bundle (Pi, Claude Code, Codex all use this)
│       ├── SKILL.md
│       ├── scripts/
│       └── requirements.txt
└── src/          ← 16-LOC pi-extension shim (skill registration only)
    └── index.ts
```

The `skills/xlsx/` directory is the canonical artifact. `src/index.ts` is a
tiny pi extension whose only job is to register `skills/` as a skill path when
the package is installed via `pi install`. Claude Code and Codex ignore it and
consume `skills/xlsx/` directly via symlinks.

Why `skills/xlsx/` (not bare `skill/`)? Pi enforces the [Agent Skills
spec](https://github.com/anthropics/skills): a skill's `SKILL.md` frontmatter
`name:` must match its parent directory name. Our `name: xlsx` therefore lives
in `xlsx/`, contained in `skills/` (the registered skill-path).

## Usage

Once the skill is on the host's skill path, any prompt like "edit this xlsx"
or "create a budget spreadsheet" will trigger the workflow:

1. `python3 <skill>/scripts/xlsx.py inspect <file>` — orient (sheets, dims, formulas)
2. Edit via `openpyxl` in a bash heredoc
3. Save via `xlsx_kit.save(wb, path)` (writes xlsx + Univer JSON + emits a snapshot event)
4. `python3 <skill>/scripts/xlsx.py validate <file>` — audit + recalc in one pass

Each operation is also importable as a python function:

```python
import sys
sys.path.insert(0, '<skill-dir>/scripts')
from inspect_xlsx import inspect
from validate_xlsx import validate

info = inspect("wb.xlsx")
report = validate("wb.xlsx")
```

See [`skills/xlsx/SKILL.md`](skills/xlsx/SKILL.md) for the full agent-facing workflow and
[`skills/xlsx/scripts/EVENTS.md`](skills/xlsx/scripts/EVENTS.md) for the optional event
wire contract.

## Optional: event wire

The skill emits structured events (`workbook_snapshot`, `script_event`) when
`XLSX_EVENTS_PATH=<file>` is exported. Off by default. Export the env var
before launching the host:

```bash
export XLSX_EVENTS_PATH=/tmp/xlsx-events.ndjson
claude   # or: codex / pi

# in another terminal:
tail -F /tmp/xlsx-events.ndjson | jq .
```

Useful for live workbook viewers, progress dashboards, or post-hoc inspection.
Full schema: [`skills/xlsx/scripts/EVENTS.md`](skills/xlsx/scripts/EVENTS.md).

## What the skill itself ships

```
skills/xlsx/
├── SKILL.md
├── README.md
├── requirements.txt
└── scripts/
    ├── xlsx.py                # unified CLI
    ├── inspect_xlsx.py        # pure function + per-script CLI
    ├── audit_xlsx.py
    ├── recalc_xlsx.py
    ├── validate_xlsx.py
    ├── extend_formula.py
    ├── doctor.py
    ├── event_emit.py
    ├── EVENTS.md
    └── xlsx_kit/__init__.py
```

Each operation script exposes a pure function plus a CLI shim, so the agent
can either invoke subprocesses (`python3 xlsx.py validate ...`) or compose
operations in a single python heredoc — both flow through bash from the
agent's perspective and work identically in all three hosts.

## Requirements

- Python 3.10+
- `openpyxl >= 3.1`, `formualizer >= 0.5.8` (in `skills/xlsx/requirements.txt`)
- For `pi install` use: `@earendil-works/pi-coding-agent` ≥ 0.74.0 (peer dependency, host-provided)

## Troubleshooting

If a script fails with `ModuleNotFoundError`:

```bash
python3 skills/xlsx/scripts/xlsx.py doctor
```

Prints the exact python interpreter the host picked up plus a copy-pasteable
`pip install` command for any missing deps. Most issues come from the host
launching from a shell whose `PATH` doesn't include the venv that has the
deps installed.

### `python3` is broken / pyexpat error (macOS Homebrew)

Recent Homebrew installs sometimes leave `python3` pointing at a freshly-built
interpreter (e.g. 3.14) whose extension modules aren't yet wired up. Symptom:

```
Library not loaded: @@HOMEBREW_PREFIX@@/.../libexpat.1.dylib
```

Workaround: invoke a known-good interpreter explicitly. The bundled scripts
have no shebang requirement — any python ≥3.10 works:

```bash
python3.13 ~/.pi/agent/git/github.com/StTronn/pi-sheets/skills/xlsx/scripts/xlsx.py doctor
python3.13 -m pip install --user --break-system-packages -r \
  ~/.pi/agent/git/github.com/StTronn/pi-sheets/skills/xlsx/requirements.txt
```

Once `python3` itself is fixed, you can drop the explicit version.

## License

MIT — see [LICENSE](LICENSE).

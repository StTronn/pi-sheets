/**
 * pi-sheets — Pi coding-agent extension that registers the bundled xlsx skill.
 *
 * Single responsibility: tell pi where to find SKILL.md so the agent can load
 * the workflow. The skill itself (SKILL.md + scripts/) is host-agnostic and
 * also works in Claude Code and Codex via plain symlinks.
 *
 * Layout: we contribute the `skills/` directory (containing `xlsx/`) as a
 * skill path. Pi's resource loader scans `<skillPath>/<name>/SKILL.md` and
 * requires `name:` in frontmatter to match the parent directory — that's why
 * the skill lives in `skills/xlsx/` not `skill/`.
 */
import type { ExtensionAPI } from "@earendil-works/pi-coding-agent";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const SKILLS_DIR = resolve(dirname(fileURLToPath(import.meta.url)), "..", "skills");

export default function piSheets(pi: ExtensionAPI) {
  pi.on("resources_discover", async () => ({ skillPaths: [SKILLS_DIR] }));
}

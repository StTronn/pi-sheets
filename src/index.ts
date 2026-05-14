/**
 * pi-sheets — Pi coding-agent extension that registers the bundled xlsx skill.
 *
 * Single responsibility: tell pi where to find SKILL.md so the agent can load
 * the workflow. The skill itself (SKILL.md + scripts/) is host-agnostic and
 * also works in Claude Code and Codex via plain symlinks.
 */
import type { ExtensionAPI } from "@earendil-works/pi-coding-agent";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const SKILL_DIR = resolve(dirname(fileURLToPath(import.meta.url)), "..", "skill");

export default function xlsxSkill(pi: ExtensionAPI) {
  pi.on("resources_discover", async () => ({ skillPaths: [SKILL_DIR] }));
}

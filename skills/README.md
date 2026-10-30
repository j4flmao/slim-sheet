# slim-sheet skill packs

Agent skills that turn the slim-sheet MCP tools into finished deliverables. Each pack follows the
Agent Skills format (`SKILL.md` with frontmatter, reference files one level deep), which Claude
(Claude Code, Claude.ai, the API) and Microsoft Copilot's Excel skills both read.

| Pack | What it produces |
|---|---|
| [building-client-reports](building-client-reports/SKILL.md) | Client-ready dashboards and report workbooks: KPI tiles with variance vs prior and plan, message titles, ranked tables, accessible colours, print-ready pages, a readiness check |

## Install

- **Claude Code**: copy the folder to `~/.claude/skills/` (all projects) or `.claude/skills/` in a
  project.
- **Claude.ai / Claude Desktop**: zip the pack folder and upload it under Settings → Capabilities →
  Skills.
- **Other agents**: paste `SKILL.md` into the system prompt and attach the reference file that
  fits the request (for example `reference/roles/cfo.md`). Small local models: start from
  `reference/small-models.md`.

The slim-sheet MCP server must be attached; see [docs/hosts.md](../docs/hosts.md).

## Quality

`tests/test_skill_pack.py` runs every role preset as written against a sample workbook, so the
presets cannot drift from the tool grammar. `evals/evals.json` lists agent-level scenarios to grade
by hand or with an eval harness.

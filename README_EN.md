# Academic Report Craft

[中文](README.md) | English

Agent Skills for writing Chinese academic documents with Claude Code or Codex: thesis proposals (开题报告), theses, project reports and competition reports.

This is a personal collection of writing and formatting habits gathered while working on a thesis proposal and thesis. It is a learning project and will keep changing.

The rules were distilled from real work: a sentence-by-sentence revision of a submitted competition report, and the writing, formatting and figure work for several thesis proposals. Each rule states its reason; see the [design rationale](skills/chinese-engineering-report-style/references/design-rationale.md) (in Chinese).

## Skills

| Skill | Scope |
| --- | --- |
| [chinese-engineering-report-style](skills/chinese-engineering-report-style/SKILL.md) | What to write and how: paragraph tasks, consistent terms, boundaries between method and results, figure references, equation layout, italic/upright symbols, simplified GB/T 7714 references |
| [typora-markdown-conventions](skills/typora-markdown-conventions/SKILL.md) | Markdown semantics for Typora and exported PDF (math vs. code markup), with an audit script |
| [word-office-automation](skills/word-office-automation/SKILL.md) | Driving Word from PowerShell: sentence-level replacement that keeps super/subscripts, converting Word equations and inline variables to MathType, line spacing, PDF checks, and re-typesetting program-generated figures in SimSun + Times New Roman |
| [visio-automation](skills/visio-automation/SKILL.md) | Editable Visio diagrams through COM with native masters and glued connectors, plus a layout standard for academic framework figures |

The first two are text-only and work anywhere. The last two need Windows and Microsoft Office (Word, MathType, Visio).

## Install

Claude Code plugin:

```text
/plugin marketplace add jxnxkx/academic-report-craft-zh
/plugin install academic-report-craft-zh
```

Or copy folders from `skills/` into `~/.claude/skills/` (Claude Code) or `~/.codex/skills/` (Codex). visio-automation keeps the upstream script path `$env:USERPROFILE\.codex\skills\visio-automation\scripts\…`; under Claude Code, replace `.codex\skills` with `.claude\skills`.

## Credits and license

- visio-automation is adapted from [renzo1031/visio-automation-skill](https://github.com/renzo1031/visio-automation-skill) (MIT). Changes are listed in [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md); the original license is kept at `skills/visio-automation/LICENSE-upstream.txt`.
- Released under the [MIT License](LICENSE).

These are writing and formatting preferences, not institutional or national requirements. A school template or supervisor's instructions take precedence.

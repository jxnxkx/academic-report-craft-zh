# Third-Party Notices

## visio-automation

`skills/visio-automation` is adapted from [renzo1031/visio-automation-skill](https://github.com/renzo1031/visio-automation-skill), commit `2dada27` (2026-06-20). The original is released under the MIT License, Copyright (c) 2026 ren zhong. The full license text is kept at `skills/visio-automation/LICENSE-upstream.txt`.

Changes in this repository:

- **New `references/academic-framework-figures.md`.** A layout standard for thesis and paper framework figures: a shared grid, uniform box, text and line sizes, labels, Z-shaped routes, and inserting into Word. It also collects Visio and Office COM pitfalls met in practice.
- **New `references/incremental-maintenance.md` and `scripts/visio_maintenance.ps1`.** Editing existing diagrams without rebuilding them: snapshots, backups, controlled PNG export, and connector endpoint reports.
- **`SKILL.md`.**
  - The description now covers thesis figures.
  - New default rules for editing versus rebuilding, and for stable object identity.
  - The reading table lists the two new references.
  - The workflow and verification steps cover incremental edits.
  - The upstream keyword list was removed.
- **`scripts/visio_helpers.ps1`.**
  - The graph model is validated before Visio starts: unique node and edge IDs, edges pointing to existing nodes, and no conflicting `-Visible`/`-Invisible`.
  - The renderer refuses to overwrite an existing preview, or an output document that is open, unless asked.
  - `-Force` copies existing outputs to a dated backup folder before rebuilding.
- **`references/graph-model-and-conversion.md` and `references/verification.md`.** Document the rules above, export-setting restoration, and endpoint checks.

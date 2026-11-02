---
name: visio-automation
description: >-
  Use when the user wants an AI agent to create, edit, convert, inspect, or repair editable Microsoft Visio diagrams with real Visio COM automation, including thesis, 开题报告 and paper figures (method frameworks, training pipelines, technical roadmaps) and tidying or straightening a figure the user has rearranged by hand.
metadata:
  upstream: "https://github.com/renzo1031/visio-automation-skill (MIT)"
  local_revision: "2026-11-02"
---

# Visio Automation

Create and modify diagrams by controlling real Microsoft Visio through COM automation. The target artifact is an editable `.vsdx` using built-in Visio masters, native dynamic connectors, glued endpoints, and ShapeSheet settings that survive manual editing.

## Default Rules

- Run user-facing Visio work in foreground visible mode by default so the user can watch the diagram appear. Use `New-InvisibleVisioApplication` only when the user explicitly asks for background/headless operation, or for narrow internal discovery/verification where no user-facing document is being drawn.
- Prefer Visio built-in stencils and `NameU` masters. Use primitive basic shapes only when no semantic master exists or the user wants a deliberately generic sketch.
- Build a lightweight Graph Model before drawing non-trivial diagrams, conversions, image reconstructions, or tasks without clear coordinates.
- Use Visio connector masters plus `GlueToPos`; do not draw loose lines for relationships that should move with shapes.
- Preserve requested or observed connector routing: orthogonal stays orthogonal, straight stays straight, curved stays curved.
- Distinguish creating/rebuilding from editing an existing document. For edits, reuse the open document, read its current native state, back up the saved state, and patch only requested fields. Do not replace unrelated manual edits from an older Graph Model.
- Identify pages and objects by stable names or explicit mappings. Repeated execution should reuse the intended objects rather than duplicate pages, shapes, connectors, or hyperlinks.
- Save the `.vsdx`, export a PNG preview when useful, then verify the resulting document before reporting completion.

## Choose What To Read

Read only the files needed for the current request:

| Task | Read |
| --- | --- |
| Choose diagram family, stencil, or master | `references/diagram-selection.md`, then `references/master-catalog.md` if needed |
| Unknown or localized master/stencil | `references/master-discovery.md` |
| Natural-language master lookup, Chinese/English aliases, repeated edits | `references/semantic-search-and-cache.md` |
| Render a planned graph, auto-layout, natural-language flow, arrow text flow, Mermaid, or draw.io | `references/graph-model-and-conversion.md` |
| Recreate screenshot, whiteboard, PDF crop, or sketch | `references/image-reconstruction.md` |
| Edit existing diagrams, add pages, preserve manual changes, resume a partial run | `references/incremental-maintenance.md` |
| Verify routing, glue, ShapeSheet cells, output files | `references/verification.md` |
| Need Visio COM, VBA, or ShapeSheet API details | `references/official-docs.md` |
| Thesis/paper framework figure, the user's figure style (grid, sizes, line weights, labels), straightening a hand-arranged figure, inserting into Word | `references/academic-framework-figures.md` |

Always load the helper script instead of retyping COM boilerplate:

```powershell
. "$env:USERPROFILE\.codex\skills\visio-automation\scripts\visio_helpers.ps1"
```

For non-Codex agents, replace the `.codex\skills` segment with that agent's skills directory.

Useful helpers include:

- Visio app/document: `New-VisibleVisioApplication`, `New-InvisibleVisioApplication`, `Get-OpenVisioDocumentByPath`
- Stencils/masters: `Open-VisioStencilReadOnly`, `Get-VisioEnvironment`, `Build-VisioMasterIndex`, `Find-VisioMasters`, `Find-SemanticVisioMaster`
- Graph rendering/conversion: `Render-VisioGraphModel`, `Invoke-VisioGraphLayout`, `Convert-NaturalLanguageFlowToVisio`, `Convert-TextFlowToVisio`, `Convert-MermaidToVisio`, `Convert-DrawIOToVisio`, `Convert-ImageReconstructionToVisio`
- Shape/connector editing: `Set-VisioTextStyle`, `Set-VisioShapeFill`, `Set-VisioShapeBounds`, `Connect-VisioShapesOrthogonal`, `Connect-VisioShapesStraight`, `Connect-VisioShapesCurved`
- Routing policy: `Resolve-VisioRoutingIntent`, `Resolve-VisioRouteDecision`, `Get-VisioRoutingDefault`, `Resolve-VisioLoopbackGlue`
- Session cache: `Get-VisioSessionCache`, `Update-VisioSessionCache`, `Add-VisioSessionRecentMaster`

For maintenance, snapshots, controlled PNG export, or endpoint checks, also load:

```powershell
. "$env:USERPROFILE\.codex\skills\visio-automation\scripts\visio_maintenance.ps1"
```

This adds `Get-VisioNativeSnapshot`, `Save-VisioDocumentBackup`, `Export-VisioPagePreview`, and `Get-VisioPageConnectorReport`; it does not replace the original renderer. The renderer accepts a single-page Graph Model. Multi-page updates use a task script as described in the maintenance reference.

## Core Workflow

1. Establish whether this is a new diagram, an explicit rebuild, or an edit of the current native document; classify the diagram family.
2. Read the relevant reference file from the table above.
3. Select native stencils/masters through the catalog, semantic search, or local discovery.
4. Normalize the diagram into a Graph Model when the structure is more than a few manually placed shapes.
5. Render or apply the targeted patch with visible Visio by default, using native masters and glued dynamic connectors.
6. Save `.vsdx`, synchronize changed model fields, export relevant pages, and check endpoint identity as well as visual routing. Report saved, exported, and verified stages separately if a run is incomplete.

## Practical Notes

- Use `pwsh` for scripts containing UTF-8 Chinese text. Windows PowerShell 5 may misparse UTF-8 files without BOM.
- Keep stencils hidden/read-only with `OpenEx(..., 66)` to avoid save prompts.
- Do not close the user's visible Visio window unless they explicitly ask.
- Avoid generating raw `.vsdx` XML for editable diagrams; it often lacks the native master and connector behavior the user cares about.
- Keep the task script in the workspace when future edits are likely.
- Do not use `Render-VisioGraphModel -Force` as an incremental editor. It rebuilds a diagram; existing output is backed up before replacement, and an open target must be edited through its current document instead.
- Preview size settings belong to the application. If temporarily changed, restore the previous settings in `finally`, including on export failure.
- In a multi-agent task, delegate source/model preparation in parallel and keep Visio mutations with one writer. User preferences such as monochrome or horizontal layout apply when requested; they are not universal defaults.

# Incremental Maintenance and Multiple Pages

Use for changes to an existing native document, adding pages, or continuing a partially completed edit. For a new single-page diagram, the ordinary renderer is sufficient.

## Current Document Is the Editing Baseline

Reuse `Get-OpenVisioDocumentByPath`; open the file only if it is not already available. Keep the user-facing application visible. Capture the current native text, geometry, links, and connector endpoints before applying a patch. `Get-VisioNativeSnapshot` provides a readable record, including grouped shapes, but is not a lossless serializer for rich text, themes, or every ShapeSheet row. Preserve those features by leaving unrequested fields untouched.

When a stored model and the open document differ, retain native changes outside the requested patch. A missing model node is not authorization to delete a shape. Do not silently replace native labels, positions, routes, styles, or hyperlinks across all pages just because a generator can be rerun.

Use `Save-VisioDocumentBackup` before each authorized edit session to save current manual changes and copy the current file plus specified related model/notes files. It requires a previously saved document and validates related files before saving. For a requested new version, save the completed document to a new path; retain the source version. A backup is recovery evidence, not permission for unrelated changes.

```powershell
. "$env:USERPROFILE\.codex\skills\visio-automation\scripts\visio_helpers.ps1"
. "$env:USERPROFILE\.codex\skills\visio-automation\scripts\visio_maintenance.ps1"
$doc = Get-OpenVisioDocumentByPath -Path $drawingPath
if ($null -eq $doc) {
  $app = New-VisibleVisioApplication
  $doc = $app.Documents.Open($drawingPath)
} else {
  $app = $doc.Application
  $app.Visible = $true
}
$backup = Save-VisioDocumentBackup -Document $doc -BackupDirectory $backupRoot -RelatedPaths @($modelPath)
$snapshot = Get-VisioNativeSnapshot -Document $doc
$snapshot | ConvertTo-Json -Depth 50 | Set-Content -LiteralPath $snapshotPath -Encoding utf8
```

## Stable Identity and Field-Level Patches

Use `(page identity, object identity)` as the key. Named pages or an explicit page mapping remain meaningful after tabs are reordered; a numeric page index alone does not. Keep a stable page `NameU` or mapping to native page ID when display names may change. Native shape `NameU` or a dedicated `User` cell can hold object identity; record the actual native name if Visio normalizes a proposed identifier.

For each requested node or edge:

- Find the current object before adding one. Update only the requested fields; create a new object only when it is absent.
- Reconnect native connector endpoints with `GlueToPos`; update the corresponding model edge too. Do not update only the drawing and leave an old `from/to` pair in the model.
- Compare a hyperlink's `Address` and `SubAddress` before adding it; distinguish external file links from internal page links.
- After rerun, check target page/object counts and duplicate identities. Deletion should be an explicit patch operation.

Set a desired alignment after calling `Set-VisioTextStyle`, because that helper sets centered paragraph and vertical alignment. Apply a Chinese-capable font when the actual text requires it, and visually inspect mathematical glyphs. Keep existing style unless its change is requested or necessary to correct the artifact.

## Multi-Page Task Model

`Render-VisioGraphModel` currently renders one page from `page`, `nodes`, and `edges`. It does not iterate a top-level `pages` array. Use a task script for multi-page maintenance rather than passing an unsupported schema.

A useful task-owned schema is:

```json
{
  "pages": [
    {
      "id": "overview",
      "nameU": "overview",
      "name": "Overall Flow",
      "width": 20,
      "height": 8,
      "nodes": [{"id": "input", "text": "Input", "x": 2, "y": 4, "width": 3, "height": 1}],
      "edges": []
    }
  ]
}
```

Coordinates use Visio inches with center positions. Object IDs are scoped to their page. Validate duplicate IDs, referenced endpoints, page dimensions, and required input files before mutating the document. Avoid moving an existing page or its objects merely to make a layout uniform.

Find or add each page by its stable identity, update requested objects, and maintain an explicit page-to-preview-path manifest. Export and verify pages by identity, not by their current tab positions. Include relevant background pages and grouped shapes in inspection; do not silently omit them when they affect the requested diagram.

## Save, Export, and Resume Are Separate Stages

An `UndoScope` can roll back Visio edits before commit. It does not make `.vsdx`, JSON, notes, PNG exports, or a `SaveAs` transactionally atomic. Do not describe it as a complete disk rollback.

Track the completed stages: native document saved, model synchronized, previews exported, structural checks passed, visual inspection performed. If native save succeeded but export failed, report that distinction. On resume, inspect the current document and available files; reconstruct missing metadata from native state where possible instead of assuming every prior-stage file exists. Take a fresh backup before further mutations.

For previews, use `Export-VisioPagePreview`. It temporarily sets raster size, restores the prior setting even when export fails, and returns file information. It does not itself prove visual correctness. Inspect dense regions at readable scale; whole-page thumbnails cannot establish that formulas and labels are legible.

For loopback connectors, retain glued endpoints and the requested route. If a label touches a box, adjust the connector text block rather than replacing the connector with loose lines. A `TxtPinY` offset can separate a label from a horizontal return path; the amount and sign depend on the actual geometry and must be checked in the preview.

For diagrams that cite documents, use accessible stable project references when saving alongside source notes is in scope. Mark unsupported or undecided content explicitly. Do not turn discussion hypotheses into verified source or experiment claims while synchronizing labels.

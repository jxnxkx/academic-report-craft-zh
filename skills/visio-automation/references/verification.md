# Verification

Use this file when you need to verify routing, glue, ShapeSheet cells, output files, or preview fidelity.

## Verification Checklist

Before claiming the diagram is complete, verify:

- The `.vsdx` file exists and has a recent timestamp.
- The document opens or is attached through ROT if already open.
- Node count and connector count are plausible.
- Connectors are `OneD` shapes.
- Default connectors should use right-angle/orthogonal routing unless the user requested straight lines.
- Image reconstruction preserves each observed connector route type: right-angle, straight, or curved.
- Straight connector requests have `ShapeRouteStyle = 2` and `ConLineRouteExt = 1`.
- Curved connector requests have `ConLineRouteExt = 2`.
- Connector endpoints remain glued; `BeginX` and `EndX` should usually contain glue formulas after connection.
- Verify the start and end separately and compare their target shape IDs with the intended model endpoints. A connection count of two alone cannot prove endpoint identity.
- For multi-page work, verify all requested pages by stable identity, including relevant grouped connectors. Do not assume tab order is unchanged.
- A preview export visually matches the requested layout closely enough.

For structural inspection, load the maintenance helpers. An endpoint-only report and a model-target report answer different questions:

```powershell
. "$env:USERPROFILE\.codex\skills\visio-automation\scripts\visio_helpers.ps1"
. "$env:USERPROFILE\.codex\skills\visio-automation\scripts\visio_maintenance.ps1"
$doc = Get-OpenVisioDocumentByPath 'E:\path\diagram.vsdx'
if ($null -eq $doc) {
  $visio = New-InvisibleVisioApplication
  $doc = $visio.Documents.Open('E:\path\diagram.vsdx')
}
$page = @($doc.Pages | Where-Object { $_.NameU -eq 'overall' }) | Select-Object -First 1
if ($null -eq $page) { throw 'Requested page not found.' }
$expected = @(
  [pscustomobject]@{ id='flow_1'; from='input'; to='process'; connector='Dynamic connector' }
)
$report = Get-VisioPageConnectorReport -Page $page -ExpectedEdges $expected
if (-not $report.allEndpointsGlued -or -not $report.expectedTargetsVerified) {
  $report | ConvertTo-Json -Depth 12
  throw 'Connector verification failed.'
}
```

Expected edge `id`, `from`, and `to` denote native `NameU` values. For an ambiguous name inside a group, supply native `fromId` / `toId` instead. The helper examines nested shapes, records each endpoint's `FromCell` and `ToSheet` ID, and reports specific missing or wrong targets. Without `ExpectedEdges`, `expectedTargetsVerified` is false; an endpoint-only inspection must not be reported as verification against a model. Master names can vary with the selected native stencil, so compare a master only when one is explicitly expected.

Use `Connect.FromCell.Name` and `Connect.ToCell.Name` for cell names. Unlike Shape/Page/Master objects, the `Cell` object's universal name is exposed as `Name`, not `NameU`; the distinction was checked in native COM and [Microsoft's Cell.Name documentation](https://learn.microsoft.com/en-us/office/vba/api/visio.cell.name). [Connect.FromCell](https://learn.microsoft.com/en-us/office/vba/api/visio.connect.fromcell) documents the begin/end cells and why some glue types produce more than two records.

The report lists one-dimensional shapes; decorative lines can be deliberately unglued. For such diagrams, use the per-connector details and check the intended edge set rather than claiming every line must be a connector.

For a deterministic PNG preview, call:

```powershell
$preview = Export-VisioPagePreview -Page $page -Path $previewPath -PixelsPerInch 144
```

Alternatively specify both `PixelWidth` and `PixelHeight`. The helper restores the application's previous raster-size settings on success and failure. It changes no color, background, or resolution settings. Check the resulting image itself and any dense crops; structural verification does not establish legible text, correct formula glyphs, uncrossed labels, or the absence of clipping.

Record document save, model synchronization, page exports, structural checks, and visual review as separate outcomes. If an invisible application was created solely for inspection, close only the document/application owned by that inspection; do not quit a reused user session.

## Practical Notes

- Use `pwsh` for scripts containing UTF-8 Chinese text. Windows PowerShell 5 may misparse UTF-8 files without BOM.
- Run user-facing Visio work in foreground visible mode by default.
- Keep stencils hidden/read-only with `OpenEx(..., 66)` to avoid save prompts.
- Avoid generating raw `.vsdx` XML for editable diagrams.

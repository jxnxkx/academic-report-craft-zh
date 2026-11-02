# Academic Framework Figures

Use this for method or framework figures that go into a thesis, a 开题报告 or a paper. Typical subjects are multi-stage training pipelines, network overviews and technical roadmaps. Read it too when tidying a figure the user has rearranged by hand.

These rules come from one author's revisions of their 开题报告 (thesis-proposal) figures. They are a house style for Chinese academic figures: follow them by default, and follow the user's own instructions when they differ.

## What a finished figure looks like

**One grid for the whole figure.**
- Every row, in every panel or stage, shares the same column edges. Stacked stages then line up vertically. At minimum, each column has one aligned edge, usually the left.
- Boxes in a row share one centre height, which is what makes horizontal connectors straight.
- A good reference is a network figure in which every row reuses the same columns, even a fine-tuning row set apart by a dashed line.

**Uniform sizes.**
- All module boxes are the same height, and boxes in a column are the same width.
- A module that feeds or receives two rows (e.g. a generator G with two inputs or two outputs) is drawn tall: two rows plus the gap. Each partner then connects with a straight line.
- Never assemble a figure from panels drawn or pasted at different scales. One such figure ended up with half its text a size smaller than the rest, and the user rejected it.

**Text sizes.**
- Box text uses one size throughout. Line labels are one step smaller, in italic. Panel titles are bold and slightly larger than box text.
- Draw at final print size and insert into Word at 100%, so the drawn point size is the printed point size. Word text width is about 14.7 cm.
- 8 pt box text was accepted. About 5 pt was rejected as unreadable on paper; that size came from a wide side-by-side layout shrunk to page width.

**Line weights.**
- Box outlines, connectors, dashed boxes and dashed lines all use one weight.
- Scale the weight to the text, roughly 1/8 of the box text size. The user's reference used 1.5 pt with 12 pt text, so 1.0 pt suits 8 pt text.
- Do not use thick lines, thick borders or bold text for emphasis; the user removed all of these. Carry emphasis in words such as "trainable" or "inference".
- Use small arrowheads (EndArrowSize 1 at 1 pt). Very small dashed chips may need 0.75 pt so the dashes don't break up.

**Connectors.**
- Lines are horizontal or vertical, with a corner only where the geometry forces one.
- Turn line jumps off.
- No label sits on a line or an arrowhead. Put it above a horizontal run or beside a vertical riser, about 0.05 to 0.1 cm clear.

**Layout.**
- If side-by-side stages would shrink the text, stack them, stage 1 above stage 2.
- Keep linked stages in one figure rather than splitting them into 图4 and 图5. The link between stages is the point of a framework figure, and the other framework figures are single figures too.
- Left-align titles such as "(1) Stage 1: …" and place each about 0.12 cm above the first row of its stage. Leave at least 0.25 cm between one stage's last row and the next title.
- Put the legend in one row at the bottom.

**Visual vocabulary.**
- The style follows CLIP (Radford et al., 2021) Fig. 1: panels titled by what the step does, straight arrows, few words, fan-in or fan-out buses.
- Small colour chips inside input boxes stand for class groups: blue for historical or seen classes, orange for new or unseen classes. A dashed orange "?" chip means unlabeled, class unknown.

**Content discipline.**
- Show the mechanism only. Leave out dimensions, step counts and hyperparameters such as τ and β.
- Formulas: only the top-level objective (e.g. L = J + αR). Make the variables italic and set Greek letters in the Latin font; the Asian fallback font renders α badly and spaces it oddly.
- Figure terms must match the body text (Semantic estimator ↔ 语义估计网络), and abbreviations used in the figure must be defined in the text (e.g. AP).

**Language in Chinese reports.**
- Overview and roadmap figures are in Chinese.
- Network and framework figures may stay in English when earlier reports from the same group do so.
- Captions are in Chinese.

## Working with the user

1. **The user places boxes; Claude straightens lines.** The user drags boxes by hand to get the compactness they want. Do not re-lay out what they placed unless asked.
2. **Read before patching.** After the user saves, read the live geometry: left, right, bottom, top and centre for every box. Report each misalignment as a number (e.g. "column 3, stage 2 is 0.05 cm right of stage 1"). The user wants to see what was off before anything moves.
3. **Straighten-only pass.** In one script:
   - snap near-misses of 0.1 cm or less to the shared column edges and row centres;
   - equalise box heights;
   - re-glue every connector and re-place the line labels;
   - export.
   Leave alone anything the user said to keep (e.g. "第二列不要动"). Report every nudge.
4. **Check the saved state.**
   - Check `Document.Saved` before exporting. If it is false, ask the user to press Ctrl+S rather than saving their session for them.
   - Confirm which variant they saved. Once, the saved file held a different variant from the one in their screenshot.
5. **Show the result and continue.** Show the exported PNG, and crop dense regions to check labels against arrowheads. The user replies with screenshots plus short Chinese notes; expect several rounds.

## Techniques

**Straight line between boxes of different heights.** Glue on the tall box at the partner's centre height. The partner side stays at 0.5.
```powershell
$rel = ($yPartnerCentre - $tallBottom) / $tallHeight
$conn.CellsU('BeginX').GlueToPos($tall, 1, $rel)
$conn.CellsU('EndX').GlueToPos($partner, 0, 0.5)
```

**Labels must not block routing.** Dynamic connectors route around 2D shapes, text labels included. A label in the path causes hooks and detours. Set `ObjType = 4` on label shapes so routes ignore them.

**Placing the middle run of a Z-route.**
1. Glue the connector.
2. Convert the target page Y to local coordinates with `XYFromPage`.
3. Write that value into Geometry1 rows 2 and 3, Y column.
4. Set `ConFixedCode = 2` so Visio keeps the route. Reset it to 0 before re-gluing later.

`ShapeRouteStyle = 5` (flowchart top-to-bottom) gives down–across–down routes.

**Measure text before choosing widths.** Add a temporary user cell with `TEXTWIDTH(TheText)` or `TEXTHEIGHT(TheText,TxtWidth)`. TEXTHEIGHT includes the margins, so set label margins to 0 before comparing.

**Chips inside boxes.** Chips are separate shapes, so move them with their box. To keep the box text above the chip strip, set `TxtHeight = Height-0.34 cm`, `TxtPinY = (Height+0.34 cm)/2` and `VerticalAlign = 1`.

**Character formatting.**
- `Characters.CharProps(i)` is an indexed property: write `$c.CharProps(2) = 2` for italic. Calling it with two arguments fails.
- To read formatting back, use `CharPropsRow(0)`, then `CellsSRC(3, row, col)`.
- Set a property on every row of the Character section. Setting row 0 alone misses runs created by earlier formatting.

**Export.**
- Call `SetRasterExportResolution(3, 600, 600, 0)`, then `Page.Export` to `.png` (it crops to the drawing extent) and to `.emf`.
- Restore the previous resolution settings in `finally`.

## Visio COM pitfalls met on this machine

These came from drawing thesis figures with this skill.

- **Plain string paths.** Pass `[string]` paths to COM. A `Join-Path` result passed to `Documents.OpenEx` makes Visio hang forever with no dialog.
- **Loading the helpers.** `visio_helpers.ps1` has no BOM. Load it with `. ([ScriptBlock]::Create([IO.File]::ReadAllText($path,[Text.Encoding]::UTF8)))`.
- **Instances.** Prefer `New-InvisibleVisioApplication` with `AlertResponse = 7`, then `Close()` and `Quit()`. A visible instance left behind blocks later runs.
- **Subscripts.** Use `Characters.CharProps(4)` (visCharacterPos; 2 = subscript). A range that reaches the end of the text, or an empty range at the end, applies to the whole text. Set tail ranges only: `[start,end)` to 2, then `[start+len,end)` to 0, the second step only if `start+len < length`.
- **Changing text.** `Characters.Text = …` throws. Replace the whole `Shape.Text`, then apply the subscripts again.
- **Duplicating pages.** `Page.Duplicate()` inserts the copy right after the source and shifts the page indexes. In a loop, look pages up by `NameU`, never by index.
- **Chinese text.** Set `Char.AsianFont` to the 宋体 font ID on every character row. Get the column from `CellsU('Char.AsianFont').Column`.

## PowerShell and Office pitfalls met here

- **Script encoding.** Windows PowerShell 5.1 needs UTF-8 with BOM for Chinese text. Writing a large script through a bash heredoc can break on quoting; write the file with the file tool, then prepend the BOM.
- **Aliases beat functions.** A helper named `R` silently ran `Invoke-History`. Avoid short names that are built-in aliases, such as `r`, `h` and `gc`.
- **Closing Word.** Call `$doc.Close([ref]0)` and `$word.Quit([ref]0)`. Plain `Quit(0)` fails in PowerShell. `New-Object -ComObject Word.Application` can attach to the Word the user has open. Record `Documents.Count` right after connecting, and quit only if it was 0 and is 0 again after closing your own document. An unconditional `Quit()` closed the user's Word on 2026-11-02. More Word recipes are in the `word-office-automation` skill.
- **Hung Office instances.** Only stop a Word or Visio process whose command line contains `/Automation -Embedding` and which this session started.
  - Killing Word leaves a resiliency flag. The next open of that path shows a hidden "上次打开…时出现严重错误" prompt and the script hangs; afterwards Word opens a repaired, unsaved copy.
  - Fix: work on a copy at a new path, then copy the result back.
  - UI Automation can read the text of a hidden dialog.
- **The user's open files.** Before overwriting a `.docx`, check for an owner file (`~$…`) and for a `WINWORD.EXE` without `/Automation`. If the user has it open, do not overwrite; ask them to close it first.

## Putting the figure into a Word report

- Find the inline picture by the caption paragraph that follows it and replace it in place.
- Insert at natural size, px / 600 × 72 pt, capped at the text width.
- Re-export the PDF and look at the pages around each figure.
- A tall figure pushed to the next page leaves a half-page gap. Report it and propose moving the figure after the following paragraph, instead of shrinking the figure and its text.

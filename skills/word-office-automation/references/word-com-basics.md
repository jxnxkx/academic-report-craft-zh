# Word 基础操作

## 打开、另存与导出

```powershell
$w = New-Object -ComObject Word.Application
$had = $w.Documents.Count          # 记下连接时已有的文档数
$w.Visible = $false; $w.DisplayAlerts = 0
$d = $w.Documents.Open([string]$src, $false, $true)   # 只读打开
$d.SaveAs2([string]$dst, 16)       # 16 = docx；.doc 用 0
$d.ExportAsFixedFormat([string]$pdf, 17)              # 17 = PDF
"pages=" + $d.ComputeStatistics(2)
$d.Close(0)
if ($had -eq 0 -and $w.Documents.Count -eq 0) { $w.Quit() }
```

- 路径一律用 `[string]` 转换后再传给 COM。`Join-Path` 返回的对象会让 Office 卡住不报错（Visio 的 `OpenEx` 实测会一直挂起）。
- 需要直接修改时先复制到新文件，再用读写方式打开副本。

## 安全

- 改动前查锁文件 `~$…`。用户开着的文件不覆盖；目标文件被占用时（`Device or resource busy`）换一个新文件名保存，并告诉用户。
- 结束时只关闭自己打开的文档。连接时 `$had` 不为 0，说明连上的是用户的 Word，不要 `Quit()`。
- 卡死的 Word 只结束命令行带 `/Automation -Embedding`、由本次会话启动的进程。强行结束会留下“上次打开时出现严重错误”的恢复标记，下次打开同一路径会被隐藏对话框挂住；改用新路径的副本。
- 用户正开着文档、又只需改几个字时，可以不经过 Word：解压 `.docx`，在 `word/document.xml` 里替换完整的 `<w:t>` 文本，再按原条目顺序重新打包。只适合替换内容恰好落在单个文本节点里的情况。

## 按句替换正文

适用于润色、统一说法等批量修改。

1. 先把正文抽成文本：遍历 `document.xml` 的段落和文字段，上标记作 `^…^`，下标记作 `_…_`，公式记作占位符。修改清单在这份文本上写。
2. 每条修改是 `{old, new}`。写入前用脚本核对：
   - `old` 在去掉标记后的文本里恰好出现一次；
   - `old` 在带标记的文本里也能原样找到，即不跨上下标，否则替换后会丢掉引用号或下标格式；
   - `old` 不超过 255 个字符（Word `Find` 的上限），过长就按句拆开。
3. 用 `Find.Execute(old, MatchCase=$true, …, Wrap=0)` 定位，再给 `Range.Text` 赋新文本。新文本继承原范围第一个字符的格式，所以范围要从普通文字开始。
4. 引用号需要挪位置时，不要把它包进替换范围：先删掉括号部分，再在引用号后面插入新句子。例如“模型（图8）[41]。模型中……”分两步改：把“模型（图8）”改为“模型”，再把“。模型中”改为“，其结构如图8所示。模型中”。
5. Word `Find` 把 `^` 当作特殊字符，`old` 里不要出现 `^`。
6. 改完再抽一次文本，确认引用号上标、变量下标仍在，页数不变或变化有原因。

## 表格和表单

- 在单元格里插入文字时，位置取段落的 `Range.End - 1`，也就是单元格结束标记之前。插在行结束标记处会报“此操作对行结尾无效”。
- 新段落会继承 `LineUnitBefore`/`LineUnitAfter`（按行设置的段前段后间距）。先把这两项设为 0，再设 `SpaceBefore`/`SpaceAfter`，否则 15.6 磅的段距一直保留，表格会被挤到下一页。
- 表单里带下划线的空白栏：用等宽的文字替换同样长度的下划线空格，下划线和整行宽度都能保留。

## 行距与段落

- 用户的规则：套用之前的模板、且模板以可读性为主时（1.25 倍行距和 20 磅固定值视觉效果相近），直接用 20 磅固定行距，避免公式让段内行距不一致（见 `chinese-engineering-report-style` 的“含行内公式的正文行距”）。
- 学校模板正文常用 1.25 倍行距（`w:line="300" w:lineRule="auto"`）。改为固定 20 磅时，在 XML 中改为 `w:line="400" w:lineRule="exact"`，跳过含 `<w:drawing>` 的插图段落和只有公式对象的独立公式段落，否则图片和多行公式会被截掉。
- 含行内公式对象的段落设 `ParagraphFormat.DisableLineHeightGrid = -1`（不对齐到文档网格），否则略高的行会占两格网格，行距忽大忽小。
- 固定行距后放大到 300 dpi 查看带下标的行，确认没有被截掉。

## 核对版面

```bash
P=/c/tools/poppler/Library/bin
$P/pdftotext -f 14 -l 14 -layout -enc UTF-8 a.pdf -     # 一页的文字
$P/pdftoppm -r 100 -png -f 14 -l 14 a.pdf out/p          # 渲染整页
$P/pdftoppm -r 300 -png -f 14 -l 14 -x 380 -y 2150 -W 1700 -H 290 a.pdf out/z   # 放大局部
```

- 页码对不上时，按“图 X”“第一行文字”逐页查找，确定内容在哪一页。
- 改动前后各导出一次 PDF，对比同一页；页数变化要找出被挤下去的段落并向用户说明。

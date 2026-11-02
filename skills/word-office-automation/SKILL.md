---
name: word-office-automation
description: 在本机用 PowerShell 驱动 Word（COM）修改 .doc/.docx：填表、按句替换正文而保留上下标、把 Word 公式和行内变量转成 MathType、调整行距、导出 PDF 并逐页核对；也用于把报告里程序生成的图重排为宋体加 Times New Roman。用户提到 Word、docx、开题报告/论文的排版或公式、MathType、“图里的字像 AI 生成的”时使用。写作取舍（写什么、怎么写）归 chinese-engineering-report-style。
---

# Word 与 MathType 自动化

在 Windows PowerShell 5.1 中通过 COM 操作 Word 和 MathType，配合 Python 和 poppler 做检查。本文件只写操作方法；正文怎么写、版式怎么取舍，见 `chinese-engineering-report-style`。

## 先做的三件事

1. **看文件是否被打开。**目录里有 `~$` 开头的锁文件，说明有人在 Word 里开着这份文件。不要覆盖，单独确认这一点，再另存新名字。
2. **在副本上改。**原文件只读；改动写进新文件名，例如 `……_图片字体已改.docx`。需要覆盖时先备份，不删除，只移动。
3. **不关别人的 Word。**`New-Object -ComObject Word.Application` 可能连上用户正在用的 Word 实例。只关闭自己打开的文档；调用 `Quit()` 前确认这个实例在打开前没有其他文档（见 [Word 基础](references/word-com-basics.md#安全)）。实际发生过：脚本末尾无条件 `Quit()`，把用户正在用的 Word 关掉了。

## 按任务读取

| 任务 | 参考文件 |
| --- | --- |
| 打开、另存、导出 PDF；按句替换正文；表格填写；行距与段落格式；核对版面 | [Word 基础](references/word-com-basics.md) |
| Word 公式或行内变量转 MathType；公式段落整理；图题中的公式 | [MathType 转换](references/mathtype-conversion.md) |
| 报告里程序生成的图要改成宋体加 Times New Roman；只有位图时重画 | [图中文字重排](references/figure-text-redraw.md) |

MathType 公式本身的版式取舍（斜线分式、自适应括号）见 `chinese-engineering-report-style` 的 `references/mathtype-formatting.md`。Visio 图见 `visio-automation`。

## 运行环境

以下是编写本Skill时的参考环境。路径不同时，在项目的 `CLAUDE.md` 或 `AGENTS.md` 里写明实际位置。


- 只有 Windows PowerShell 5.1，没有 `pwsh`。含中文的 `.ps1` 必须是带 BOM 的 UTF-8：用文件工具写好后执行 `printf '\xef\xbb\xbf' | cat - a.ps1 > ab.ps1`。
- Python 3.11 以上，需要 `matplotlib`、`numpy`、`Pillow`。Windows 上 PATH 里的 `python` 可能是 WindowsApps 占位程序，会静默不运行，此时用完整路径，例如 `%LOCALAPPDATA%\Programs\Python\Python313\python.exe`。输出中文时设 `PYTHONIOENCODING=utf-8`。
- poppler（`pdftotext`、`pdftoppm`），参考环境装在 `C:\tools\poppler\Library\bin`，用于抽取文字和把 PDF 页面渲染成图片。
- MathType（实测 2025 版），Word 加载项为 `MathType Commands 2016.dotm`，COM 启动的 Word 也会加载。
- 字体：`C:\Windows\Fonts\simsun.ttc`（宋体）、`times.ttf`/`timesi.ttf`（Times New Roman 正体、斜体）。

## 交付前核对

- 导出 PDF，比较改动前后的页数；用 `pdftotext` 确认关键句，用 `pdftoppm` 渲染改动所在页并亲眼看过。
- 抽出 `word/document.xml` 检查上标引用号、下标变量是否还在，被替换的内容是否已无残留。
- 向用户说明：新文件位置、改了几处、哪几处没改及原因、页数变化。

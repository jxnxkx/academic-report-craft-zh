# Academic Report Craft

中文 | [English](README_EN.md)

这是一组写中文学术报告用的 Agent Skills，适用于 Claude Code 和 Codex，可以用来写开题报告、学位论文、项目报告和竞赛报告。

这是我写开题报告和学位论文时整理的写作与排版习惯，算是一份学习成果，会随使用持续更新。

它们来自真实的写作过程：一份竞赛项目报告的逐句修改、几份开题报告的写作、排版和改图。每条规则都写了理由，规则来由见 [设计依据](skills/chinese-engineering-report-style/references/design-rationale.md)。

## 包含的 Skill

| Skill | 管什么 | 典型用法 |
| --- | --- | --- |
| [chinese-engineering-report-style](skills/chinese-engineering-report-style/SKILL.md) | 写什么、怎么写：段落任务、术语统一、方法与实验的叙述边界、图表引用、公式排法、符号正斜体、GB/T 7714 简化著录 | “按这个风格逐句润色第三节”“把实验描述改得像人话” |
| [typora-markdown-conventions](skills/typora-markdown-conventions/SKILL.md) | Markdown 在 Typora 和导出的 PDF 中的写法：数学用 `$…$`，代码标识用反引号，附检查脚本 | 写给导师看、要打印的讨论稿 |
| [word-office-automation](skills/word-office-automation/SKILL.md) | 用 PowerShell 驱动 Word：按句替换而保留上下标、Word 公式和行内变量转 MathType、行距、导出 PDF 逐页核对；把程序生成的图重排为宋体加 Times New Roman | “全文公式改成 MathType”“图里的字一看就是 AI 生成的，改掉” |
| [visio-automation](skills/visio-automation/SKILL.md) | 用 Visio COM 画可编辑的 `.vsdx`：原生模具、粘连连接线；论文方法框架图的版式规范 | 开题报告的技术路线图、方法框架图 |

前两个只依赖文字说明，任何平台都能用。后两个需要 Windows 和 Microsoft Office（Word、MathType、Visio）。

## 一些默认规则

- 引用图写“如图 8 所示”，不在句中夹“（图 8）”；引用号紧跟它的出处。
- 数值范围写“至”，表格外不用破折号连接概念。
- 引出独立公式的句子以冒号结束；两行公式尽量拆开，递进关系的两式之间用“其中”说明推导。
- 变量斜体；数字、单位、缩写、说明性下标正体；拉丁和希腊字母用 Times New Roman。
- 正文含行内公式、模板以可读性为主时，用 20 磅固定行距，避免行距忽大忽小。
- 程序画的图交付前改为宋体加 Times New Roman、小五号，数据不动。
- 探索性的工程细节（文件名、排障过程）不进正文。

前后对照的例子见 [examples/before-after.md](examples/before-after.md)。

## 安装

**Claude Code（插件方式）：**

```text
/plugin marketplace add jxnxkx/academic-report-craft-zh
/plugin install academic-report-craft-zh
```

**手动复制：**把 `skills/` 下需要的文件夹复制到：

- Claude Code：`~/.claude/skills/`（全局）或项目内的 `.claude/skills/`
- Codex：`~/.codex/skills/`

visio-automation 的说明里，脚本路径沿用上游写法 `$env:USERPROFILE\.codex\skills\visio-automation\scripts\…`。装在 Claude Code 下时，把 `.codex\skills` 换成 `.claude\skills`。

## 使用前提

- word-office-automation：Windows PowerShell 5.1、Microsoft Word、MathType（转换公式时）、Python 3.11 以上（`matplotlib`、`numpy`、`Pillow`）、poppler（核对 PDF 时）。
- visio-automation：Microsoft Visio（桌面版）。
- typora-markdown-conventions 的检查脚本：Python 3。

## 致谢与许可

- visio-automation 改编自 [renzo1031/visio-automation-skill](https://github.com/renzo1031/visio-automation-skill)（MIT）。改动内容见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)，原许可证见 `skills/visio-automation/LICENSE-upstream.txt`。
- 本仓库以 [MIT 许可证](LICENSE) 发布。

这些规则反映的是一种写作和排版偏好，不是学校或标准的强制要求。所在单位的模板和导师的要求优先。

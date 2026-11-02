---
name: typora-markdown-conventions
description: Standardize Markdown semantics and rendering for Chinese technical reports in Typora and exported PDF. Use when drafting, rewriting, formatting, or auditing `.md` reports containing mathematical variables, inline or display equations, code identifiers, JSON fields, commands, paths, tables, figures, or mixed Chinese-English technical prose, especially when consistent rendering across agents is required.
---

# Typora Markdown Conventions

Apply one semantic rule throughout: use math markup for mathematical meaning and code markup for implementation identifiers.

## Choose the markup

| Content | Markdown | Example |
|---|---|---|
| Mathematical variable or indexed symbol in prose | `$...$` | `$H_t$`, `$b_{t+1}$`, `$\theta$` |
| Short mathematical relation in prose | `$...$` | `$P(y\mid H_t)>\tau$` |
| Standalone equation or derivation | `$$...$$` on separate lines | See display formula rule below |
| Literal field, key, function, class, command, path, or filename | backticks | `H_t`, `update_state()`, `config.yaml` |
| Multi-line source code or data | fenced code block with language | `python`, `json`, `yaml`, `bash` |
| Product, model, framework, algorithm, or general technical term | plain text | LangGraph, LoRA, reducer, Persona |
| Emphasis in prose | `**...**` sparingly | `**核心约束**` |

The same token may use different markup by meaning:

- Write “历史状态记为 $H_t$” when discussing a mathematical state variable.
- Write “JSON 字段 `H_t`” when naming the literal implementation key.
- Write “由 reducer 统一提交” when reducer is a general architectural role.
- Write “调用 `reduce_state()`” when referring to an exact function.

## Format equations

Use inline math only when the expression belongs grammatically to the sentence:

```markdown
第 $t$ 轮的历史状态记为 $H_t$，更新后得到 $H_{t+1}$。
```

Use display math for an equation that deserves its own line:

```markdown
系统状态定义为：

$$
S_t=\{H_t,E_t,b_t,C_t,R_t,B_t,Q_t,Slots_t,Trace_t\}
$$

其中，$H_t$ 表示历史对话。
```

Keep Chinese punctuation outside math delimiters. Use braces for multi-character subscripts, such as `$Slots_{t+1}$`. Do not write raw underscore variables such as `H_t` in ordinary prose because Markdown renderers may interpret them inconsistently.

## Format code

Use backticks only for literal machine-facing text:

```markdown
系统从 `config.yaml` 读取 `risk_threshold`，随后调用 `update_state()`。
```

Use fenced blocks for executable code or structured data, and always specify the language:

````markdown
```json
{
  "H_t": [],
  "risk_threshold": 0.7
}
```
````

Do not use backticks merely to emphasize English words, abbreviations, mathematical variables, or named concepts.

## Preserve report typography

- Separate headings, paragraphs, lists, tables, figures, code blocks, and display equations with blank lines.
- Do not insert leading full-width spaces or `&emsp;` for paragraph indentation; the Typora theme supplies the two-character first-line indent.
- Do not add spaces around inline math to repair visual gaps. Treat unusually large gaps as a theme justification issue, not a Markdown-content issue.
- Do not use HTML layout tags, fixed heights, manual line breaks, or repeated `<br>` elements for pagination.
- Use normal Markdown headings. Do not bold an entire heading inside the heading marker.
- Keep figure and table captions as ordinary text unless the document defines a separate caption convention.
- Use plain ASCII spaces around embedded English terms where Chinese readability requires them; do not add spaces inside math delimiters.

## Audit before delivery

1. Search for backticked tokens resembling indexed variables, such as `` `H_t` ``. Convert them to math unless the sentence explicitly names a field or key.
2. Search for raw underscore tokens outside math and code.
3. Confirm every display equation has paired `$$` delimiters on separate lines.
4. Confirm every fenced code block closes and declares a language.
5. Inspect ambiguous terms by semantics rather than applying a global replacement.
6. When editing a file, preserve unrelated author formatting and report which ambiguous tokens were intentionally left as code.

Run the optional read-only audit:

```bash
python scripts/audit_markdown.py path/to/report.md
```

Treat audit findings as review candidates, not automatic errors.

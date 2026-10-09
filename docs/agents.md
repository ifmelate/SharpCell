# Using SharpCell from an AI coding agent

How to give Claude Code, Cursor, Copilot or another coding agent accurate knowledge of SharpCell.

## Point the agent at llms.txt

The documentation is published as plain Markdown next to the HTML. Give your agent one of these:

- [llms.txt](llms.txt): a short index of every page with one line each.
- [llms-full.txt](llms-full.txt): the guides and the API reference in one file, for agents that read a whole document.
- Any page with `.md` instead of `.html` in its address, such as [the API reference](api/index.md).

## Add the rules to your project

Paste this into the file your agent reads: `AGENTS.md`, `CLAUDE.md`, `.cursor/rules/sharpcell.mdc` or `.github/copilot-instructions.md`.

```markdown
## SharpCell (Excel formula engine)

- Packages: `SharpCell` (engine) and `SharpCell.Xlsx` (reads .xlsx). Namespaces `SharpCell`, `SharpCell.Xlsx`.
- Load: `Workbook workbook = XlsxReader.Load(path);` Build in code: `new Workbook()`, `AddSheet`, `sheet["A1"].Value`, `sheet["A2"].Formula = "=A1*3"`.
- Call `workbook.Recalculate()` after changing values or formulas. Values do not update on assignment.
- `workbook.Evaluate("=SUM(1,2)")` evaluates a formula that lives in no cell.
- A `CellValue` has a `Kind`: Number, Text, Boolean, Error, Array, Empty, Missing, Lambda. Check `Kind` before `AsNumber()`, `AsText()` and the other `As` methods; they throw on the wrong kind.
- Excel errors such as #DIV/0! are values (`Kind == Error`), not exceptions. A formula that cannot be parsed throws `FormulaParseException`.
- Formulas use Excel's file syntax: English function names, commas between arguments, dot decimals; the leading `=` is optional.
- Dates are numbers (serial days). Text-to-number conversion uses `workbook.Culture`, invariant by default.
- A `Workbook` is not thread-safe; use one per thread.
- Not supported: writing files, .xls/.xlsm/.xlsb, iterative calculation, localized function names.
- Before relying on a function, check its status in compatibility.json: https://ifmelate.github.io/SharpCell/compatibility.json
- Docs for agents: https://ifmelate.github.io/SharpCell/llms.txt
```


## Check a function before using it

`compatibility.json` lists every function with its status and how many Excel-calculated cells it matched:

```bash
curl -s https://ifmelate.github.io/SharpCell/compatibility.json | jq '.functions[] | select(.name == "XLOOKUP")'
```

```json
{"name": "XLOOKUP", "status": "known deviation", "deviation": "Match mode 3 (regular expressions) uses .NET regular expressions, whose syntax differs from Excel's PCRE2 in details.", "cases": 437, "passed": 437}
```

A function with status `not implemented` evaluates to `#NAME?`.

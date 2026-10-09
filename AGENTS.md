# Working on SharpCell

Instructions for AI agents and people changing this repository. To *use* SharpCell, read [docs/agents.md](docs/agents.md).

## Layout

- `src/SharpCell` — the engine: values, parser, evaluator, functions. No dependencies.
- `src/SharpCell.Xlsx` — reads .xlsx into the engine's workbook.
- `tests/SharpCell.Tests` — xUnit tests, including the Excel corpus in `tests/corpus`.
- `tools/SharpCell.Conformance` — runs the corpus and writes `docs/compatibility.{md,json}`.
- `tools/SharpCell.Docs` — fills code snippets into Markdown, writes `docs/api`, `docs/llms*.txt`, renders the site.
- `samples/SharpCell.Sample` — the console sample shown in the getting started guide.

## Commands

```bash
dotnet build
dotnet test                                                            # net8.0 and net10.0
dotnet run --project tools/SharpCell.Conformance -f net10.0             # after changing functions
dotnet run --project tools/SharpCell.Docs -f net10.0 -- generate        # after changing docs, snippets or public API
dotnet run --project tools/SharpCell.Docs -f net10.0 -- check
```

## Rules

- Test first: a failing test, then the code, then the commit.
- Every public API change updates `PublicAPI.Unshipped.txt` and has an XML comment; the build fails otherwise.
- The engine reads no `CultureInfo.CurrentCulture`; culture comes from the workbook.
- A function that differs from Excel is registered with `KnownDeviation` and a one-sentence explanation.
- Only values calculated by Microsoft Excel count as expected results.
- Code examples in docs live in tested `.cs` files between `// snippet: name` and `// end-snippet`; never edit the body of a `csharp snippet=` block by hand.
- `docs/api`, `docs/llms.txt`, `docs/llms-full.txt` and `docs/compatibility.*` are generated.
- Code, comments, docs and commit messages are in English. Commits follow `feat:`, `fix:`, `docs:`, `test:`, `refactor:`, `build:`.

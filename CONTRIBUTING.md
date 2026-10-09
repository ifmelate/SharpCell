# Contributing

Thank you for helping. SharpCell is kept small and correct; contributions that keep it that way
are welcome.

## Bugs

Open an issue with the bug form. It needs a small .xlsx that shows the problem, the cell and its
formula, and the value Microsoft Excel calculates. Values from LibreOffice or other programs are
not a reference. An issue without a reproducing file is closed.

## Pull requests

- Start with a failing test; a new function comes with test cases for its results and errors.
- A public API change updates `PublicAPI.Unshipped.txt` and has XML documentation.
- A function that differs from Excel is registered with `KnownDeviation` and a one-sentence reason.
- Regenerate what the change affects: the compatibility report after function changes, the docs
  after documentation or public API changes. The commands are in [AGENTS.md](AGENTS.md).
- The engine takes no dependencies and never reads `CultureInfo.CurrentCulture`.
- Code, comments and docs are in English.

By contributing you agree that your contribution is licensed under the MIT license.

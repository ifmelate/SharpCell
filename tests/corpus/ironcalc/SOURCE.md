# IronCalc test corpus

Workbooks calculated by Microsoft Excel, copied from the IronCalc project and used here as
reference values: SharpCell loads each file, recalculates it and compares every formula result
with the value Excel saved.

- Source: https://github.com/ironcalc/IronCalc
- Commit: `6a1d35a6bed741a44121360bb874c79c09a54acf` (2026-10-08)
- License: MIT or Apache-2.0, at your option (`LICENSE-MIT`, `LICENSE-Apache-2.0`).
  Copyright (c) 2023 EqualTo GmbH, 2023 Nicolás Hatcher.

## What was copied

| Here | Upstream |
|---|---|
| `calc_tests/` | `xlsx/tests/calc_tests/` |
| `docs/` | `xlsx/tests/docs/` |
| `templates/` | `xlsx/tests/templates/` |
| `DynamicArrays.xlsx`, `dynamic_arrays.xlsx` | `xlsx/tests/` |
| `tables.xlsx` | `xlsx/tests/calc_test_no_export/` |

Files were not modified.

## Left out

Only values calculated by Excel count as reference. These upstream files were saved by other
applications (per `docProps/app.xml` and the XML namespaces used), so their cached values are not
Excel's:

- `calc_tests/LOOKUP_AND_REFERENCE/XMATCH_arrays.xlsx` (IronCalc)
- `calc_tests/array_in_scalar_repro.xlsx` (IronCalc)
- `calc_tests/STATISTICAL/GEOMEAN.xlsx` (Google Sheets export)
- `docs/CHOOSE.xlsx` (IronCalc)

## Conventions

Some files carry a `METADATA` sheet, following IronCalc's test runner:

- a number in `A1` is the relative tolerance for comparing numbers in that file;
- a row with `NOW` in column A holds, in column B, the `NOW()` value Excel saved; the clock is set
  to it so date and time functions can be compared exactly.

`../overrides.json` states, with a reason for each, the locale Excel calculated a file in (a file
does not record it) and wider tolerances for cells where Excel's own result is only as accurate as
its solver. The compatibility report lists them and counts those cells apart.

These files are untrusted input: they are only parsed, never opened in Excel.

# Security

SharpCell parses formulas and reads .xlsx files, often from untrusted sources. A file or formula
that makes it crash the process, hang, use unbounded memory or read outside the workbook is a
security issue.

## Reporting

Report privately through GitHub: **Security → Report a vulnerability** on
https://github.com/ifmelate/SharpCell. Do not open a public issue. Attach the smallest file or
formula that shows the problem, and do not send confidential workbooks.

The project has one maintainer working in spare time. Expect an answer within two weeks; a fix
ships in a new release of the latest minor version.

## Supported versions

Only the latest released version gets fixes.

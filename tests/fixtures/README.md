# Workbook fixtures

- `excel/blank-excel.xlsx` was created with Microsoft Excel and then sanitized by
  `scripts/sanitize_xlsx_fixture.py`. The sanitizer removes Office author metadata,
  unreferenced trash data, and opaque printer settings before the file is committed.
- A corresponding WPS-authored fixture is still required before WPS compatibility can
  be treated as verified.

Fixtures must contain only synthetic data. Never copy a workbook from a real project.

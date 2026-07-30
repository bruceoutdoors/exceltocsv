# /// script
# dependencies = ["xlwt"]
# ///
"""Generate tests/fixtures/simple.xls (legacy XLS format).

Run from the repository root:
    uv run tools/generate_fixtures.py

The data written here must match tests/fixtures/simple.xlsx so that
tests/expected/simple_xls.csv and tests/expected/simple.csv are identical.
"""

import os

import xlwt

os.makedirs("tests/fixtures", exist_ok=True)

wb = xlwt.Workbook()
ws = wb.add_sheet("Sheet1")

ws.write(0, 0, "Name")
ws.write(0, 1, "Amount")
ws.write(0, 2, "Active")

ws.write(1, 0, "Alice")
ws.write(1, 1, 100)
ws.write(1, 2, "true")

ws.write(2, 0, "Bob")
ws.write(2, 1, 250.5)
ws.write(2, 2, "false")

wb.save("tests/fixtures/simple.xls")
print("Written: tests/fixtures/simple.xls")

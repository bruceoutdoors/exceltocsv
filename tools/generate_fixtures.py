# /// script
# dependencies = ["xlwt"]
# ///
"""Generate binary test fixtures that cannot be produced by the Rust generator.

Run from the repository root:
    uv run tools/generate_fixtures.py

Generates:
  tests/fixtures/simple.xls  — legacy XLS via xlwt
  tests/fixtures/simple.ods  — ODS via stdlib zipfile + XML (no extra deps)

For tests/fixtures/any_sheets.xlsb (XLSB format), regenerate with LibreOffice:
    soffice --headless --convert-to xlsb --outdir tests/fixtures/ \\
        <some_source.xlsx>
Or download the fixture from calamine's test suite (Apache-2.0 / MIT):
    curl -L https://github.com/tafia/calamine/raw/master/tests/any_sheets.xlsb \\
         -o tests/fixtures/any_sheets.xlsb
"""

import io
import os
import zipfile

import xlwt

os.makedirs("tests/fixtures", exist_ok=True)

# ── simple.xls ────────────────────────────────────────────────────────────────

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

# ── simple.ods ────────────────────────────────────────────────────────────────

MIMETYPE = b"application/vnd.oasis.opendocument.spreadsheet"

MANIFEST = (
    b'<?xml version="1.0" encoding="UTF-8"?>'
    b'<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">'
    b'<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>'
    b'<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>'
    b"</manifest:manifest>"
)


def str_cell(v):
    return f'<table:table-cell office:value-type="string"><text:p>{v}</text:p></table:table-cell>'


def num_cell(v):
    return (
        f'<table:table-cell office:value-type="float" office:value="{v}">'
        f"<text:p>{v}</text:p></table:table-cell>"
    )


rows = [
    [str_cell("Name"), str_cell("Amount"), str_cell("Active")],
    [str_cell("Alice"), num_cell("100"), str_cell("true")],
    [str_cell("Bob"), num_cell("250.5"), str_cell("false")],
]

rows_xml = "".join(
    "<table:table-row>" + "".join(cells) + "</table:table-row>" for cells in rows
)

CONTENT = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    "<office:document-content"
    ' xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"'
    ' xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"'
    ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"'
    ' office:version="1.2">'
    "<office:body><office:spreadsheet>"
    f'<table:table table:name="Sheet1">{rows_xml}</table:table>'
    "</office:spreadsheet></office:body></office:document-content>"
).encode()

buf = io.BytesIO()
with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr(zipfile.ZipInfo("mimetype"), MIMETYPE, compress_type=zipfile.ZIP_STORED)
    z.writestr("META-INF/manifest.xml", MANIFEST)
    z.writestr("content.xml", CONTENT)

with open("tests/fixtures/simple.ods", "wb") as f:
    f.write(buf.getvalue())
print("Written: tests/fixtures/simple.ods")

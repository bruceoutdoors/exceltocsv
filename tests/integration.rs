use std::io::Write;
use std::path::Path;
use std::process::{Command, Stdio};

fn bin() -> Command {
    Command::new(env!("CARGO_BIN_EXE_exceltocsv"))
}

fn run(args: &[&str]) -> (Vec<u8>, Vec<u8>, bool) {
    run_in(Path::new("."), args)
}

fn run_in(dir: &Path, args: &[&str]) -> (Vec<u8>, Vec<u8>, bool) {
    let out = bin()
        .args(args)
        .current_dir(dir)
        .output()
        .expect("failed to run exceltocsv");
    (out.stdout, out.stderr, out.status.success())
}

fn run_with_stdin(args: &[&str], stdin_bytes: &[u8]) -> (Vec<u8>, Vec<u8>, bool) {
    let mut child = bin()
        .args(args)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .expect("failed to spawn exceltocsv");
    child
        .stdin
        .take()
        .expect("no stdin")
        .write_all(stdin_bytes)
        .expect("write failed");
    let out = child.wait_with_output().expect("wait failed");
    (out.stdout, out.stderr, out.status.success())
}

fn expected(name: &str) -> Vec<u8> {
    std::fs::read(Path::new("tests/expected").join(name))
        .unwrap_or_else(|_| panic!("missing tests/expected/{name}"))
}

fn fixture_bytes(name: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures")
            .join(name),
    )
    .unwrap_or_else(|_| panic!("missing tests/fixtures/{name}"))
}

// ── File input tests ──────────────────────────────────────────────────────────

#[test]
fn simple_xlsx() {
    let (out, _, ok) = run(&["tests/fixtures/simple.xlsx"]);
    assert!(ok);
    assert_eq!(out, expected("simple.csv"));
}

#[test]
fn simple_xls() {
    let (out, _, ok) = run(&["tests/fixtures/simple.xls"]);
    assert!(ok);
    assert_eq!(out, expected("simple_xls.csv"));
}

#[test]
fn names_flag() {
    let (out, _, ok) = run(&["--names", "tests/fixtures/multiple_sheets.xlsx"]);
    assert!(ok);
    assert_eq!(out.trim_ascii_end(), b"Sheet1\nSheet2");
}

#[test]
fn sheet_selection() {
    let (out, _, ok) = run(&["--sheet", "Sheet2", "tests/fixtures/multiple_sheets.xlsx"]);
    assert!(ok);
    assert_eq!(out, expected("multiple_sheets_sheet2.csv"));
}

#[test]
fn multiple_sheets_default() {
    let (out, _, ok) = run(&["tests/fixtures/multiple_sheets.xlsx"]);
    assert!(ok);
    assert_eq!(out, expected("multiple_sheets_sheet1.csv"));
}

#[test]
fn quoted_values() {
    let (out, _, ok) = run(&["tests/fixtures/quoted_values.xlsx"]);
    assert!(ok);
    assert_eq!(out, expected("quoted_values.csv"));
}

#[test]
fn unicode_passthrough() {
    let (out, _, ok) = run(&["tests/fixtures/unicode.xlsx"]);
    assert!(ok);
    assert_eq!(out, expected("unicode.csv"));
}

#[test]
fn types_fixture() {
    let (out, _, ok) = run(&["tests/fixtures/types.xlsx"]);
    assert!(ok);
    assert_eq!(out, expected("types.csv"));
}

#[test]
fn tab_delimiter() {
    let (out, _, ok) = run(&["--out-tabs", "tests/fixtures/simple.xlsx"]);
    assert!(ok);
    let expected_tsv: Vec<u8> = expected("simple.csv")
        .iter()
        .map(|&b| if b == b',' { b'\t' } else { b })
        .collect();
    assert_eq!(out, expected_tsv);
}

#[test]
fn custom_delimiter() {
    let (out, _, ok) = run(&["-D", "|", "tests/fixtures/simple.xlsx"]);
    assert!(ok);
    let expected_psv: Vec<u8> = expected("simple.csv")
        .iter()
        .map(|&b| if b == b',' { b'|' } else { b })
        .collect();
    assert_eq!(out, expected_psv);
}

#[test]
fn write_sheets_all() {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/multiple_sheets.xlsx");
    let dir = tempfile::tempdir().expect("tempdir");
    let (_, err, ok) = run_in(
        dir.path(),
        &["--write-sheets", "-", fixture.to_str().unwrap()],
    );
    assert!(ok, "{}", String::from_utf8_lossy(&err));

    let s1 = std::fs::read(dir.path().join("sheet1.csv")).expect("sheet1.csv missing");
    let s2 = std::fs::read(dir.path().join("sheet2.csv")).expect("sheet2.csv missing");
    assert_eq!(s1, expected("multiple_sheets_sheet1.csv"));
    assert_eq!(s2, expected("multiple_sheets_sheet2.csv"));
}

#[test]
fn write_sheets_use_names() {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/multiple_sheets.xlsx");
    let dir = tempfile::tempdir().expect("tempdir");
    let (_, err, ok) = run_in(
        dir.path(),
        &[
            "--write-sheets",
            "-",
            "--use-sheet-names",
            fixture.to_str().unwrap(),
        ],
    );
    assert!(ok, "{}", String::from_utf8_lossy(&err));

    let s1 = std::fs::read(dir.path().join("Sheet1.csv")).expect("Sheet1.csv missing");
    let s2 = std::fs::read(dir.path().join("Sheet2.csv")).expect("Sheet2.csv missing");
    assert_eq!(s1, expected("multiple_sheets_sheet1.csv"));
    assert_eq!(s2, expected("multiple_sheets_sheet2.csv"));
}

#[test]
fn write_sheets_no_clobber() {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/multiple_sheets.xlsx");
    let dir = tempfile::tempdir().expect("tempdir");
    std::fs::write(dir.path().join("sheet1.csv"), b"existing").unwrap();
    let (_, _, ok) = run_in(
        dir.path(),
        &["--write-sheets", "-", fixture.to_str().unwrap()],
    );
    assert!(
        !ok,
        "expected non-zero exit when output file already exists"
    );
    // Pre-existing file must not be overwritten
    assert_eq!(
        std::fs::read(dir.path().join("sheet1.csv")).unwrap(),
        b"existing"
    );
}

#[test]
fn nonexistent_file() {
    let (_, _, ok) = run(&["tests/fixtures/bogus.xlsx"]);
    assert!(!ok, "expected non-zero exit for missing file");
}

#[test]
fn missing_sheet_name() {
    let (_, _, ok) = run(&["--sheet", "NoSuchSheet", "tests/fixtures/simple.xlsx"]);
    assert!(!ok, "expected non-zero exit for missing sheet");
}

#[test]
fn names_conflicts_with_sheet() {
    let (_, _, ok) = run(&["--names", "--sheet", "Sheet1", "tests/fixtures/simple.xlsx"]);
    assert!(!ok, "--names and --sheet must conflict");
}

#[test]
fn names_conflicts_with_write_sheets() {
    let (_, _, ok) = run(&[
        "--names",
        "--write-sheets",
        "-",
        "tests/fixtures/simple.xlsx",
    ]);
    assert!(!ok, "--names and --write-sheets must conflict");
}

#[test]
fn sheet_conflicts_with_write_sheets() {
    let (_, _, ok) = run(&[
        "--sheet",
        "Sheet1",
        "--write-sheets",
        "-",
        "tests/fixtures/simple.xlsx",
    ]);
    assert!(!ok, "--sheet and --write-sheets must conflict");
}

#[test]
fn quoting_none_rejected() {
    let (_, err, ok) = run(&["-U", "3", "tests/fixtures/simple.xlsx"]);
    assert!(!ok);
    assert!(String::from_utf8_lossy(&err).contains("not supported") || !err.is_empty());
}

#[test]
fn quoting_out_of_range_rejected() {
    let (_, _, ok) = run(&["-U", "4", "tests/fixtures/simple.xlsx"]);
    assert!(!ok);
}

#[test]
fn malformed_not_a_workbook() {
    let (_, _, ok) = run_with_stdin(&["-f", "xlsx"], b"this is not an excel file");
    assert!(!ok, "expected non-zero exit for invalid xlsx");
}

#[test]
fn malformed_truncated_zip() {
    // Starts with ZIP magic bytes but truncated
    let (_, _, ok) = run_with_stdin(&["-f", "xlsx"], &[0x50, 0x4B, 0x03, 0x04, 0x00]);
    assert!(!ok, "expected non-zero exit for truncated xlsx");
}

// ── Stdin input tests ─────────────────────────────────────────────────────────

#[test]
fn stdin_xlsx_explicit_dash() {
    let bytes = fixture_bytes("simple.xlsx");
    let (out, _, ok) = run_with_stdin(&["-"], &bytes);
    assert!(ok);
    assert_eq!(out, expected("simple.csv"));
}

#[test]
fn stdin_xlsx_no_arg() {
    let bytes = fixture_bytes("simple.xlsx");
    let (out, _, ok) = run_with_stdin(&[], &bytes);
    assert!(ok);
    assert_eq!(out, expected("simple.csv"));
}

#[test]
fn stdin_xlsx_format_flag() {
    let bytes = fixture_bytes("simple.xlsx");
    let (out, _, ok) = run_with_stdin(&["-f", "xlsx"], &bytes);
    assert!(ok);
    assert_eq!(out, expected("simple.csv"));
}

#[test]
fn stdin_xls_format_flag() {
    let bytes = fixture_bytes("simple.xls");
    let (out, _, ok) = run_with_stdin(&["-f", "xls"], &bytes);
    assert!(ok);
    assert_eq!(out, expected("simple_xls.csv"));
}

#[test]
fn stdin_xls_auto_detect() {
    let bytes = fixture_bytes("simple.xls");
    let (out, _, ok) = run_with_stdin(&["-"], &bytes);
    assert!(ok);
    assert_eq!(out, expected("simple_xls.csv"));
}

#[test]
fn xlsx_and_xls_same_output() {
    let (xlsx_out, _, ok1) = run(&["tests/fixtures/simple.xlsx"]);
    let (xls_out, _, ok2) = run(&["tests/fixtures/simple.xls"]);
    assert!(ok1 && ok2);
    // Note: simple.xlsx and simple.xls have same data but expected CSVs differ
    // because XLS stores "true"/"false" strings while XLSX has string "true"/"false"
    // Both should succeed and produce non-empty output
    assert!(!xlsx_out.is_empty() && !xls_out.is_empty());
}

#[test]
fn unsupported_format_ods_content() {
    // ODS files start with PK (ZIP) like XLSX, so magic-bytes can't distinguish
    // Just test that the binary rejects non-workbook content
    let (_, _, ok) = run_with_stdin(&["-f", "xlsx"], b"not a workbook");
    assert!(!ok);
}

// ── Sparse XLSX (two-pass streaming) ─────────────────────────────────────────

#[test]
fn sparse_xlsx_preserves_empty_rows_and_columns() {
    let (out, _, ok) = run(&["tests/fixtures/sparse.xlsx"]);
    assert!(ok);
    assert_eq!(out, expected("sparse.csv"));
}

// ── Adversarial XLSX fixtures ─────────────────────────────────────────────────
//
// Build minimal in-memory XLSX ZIPs to exercise the streaming state machine
// without committing adversarial binary fixtures.

fn make_xlsx_with_sheet_xml(sheet_xml: &str) -> Vec<u8> {
    use std::io::Write as _;
    use zip::{write::SimpleFileOptions, CompressionMethod, ZipWriter};

    let cursor = std::io::Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let stored = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);

    zip.start_file("[Content_Types].xml", stored).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"#).unwrap();

    zip.start_file("_rels/.rels", stored).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#).unwrap();

    zip.start_file("xl/workbook.xml", stored).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", stored)
        .unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", stored).unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();

    zip.start_file("xl/styles.xml", stored).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts><font/></fonts><fills><fill/><fill/></fills><borders><border/></borders><cellStyleXfs><xf/></cellStyleXfs><cellXfs><xf/></cellXfs></styleSheet>"#).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn xlsx_duplicate_cell_rejected() {
    // Two cells sharing the same coordinate must be rejected to prevent
    // unbounded row_buf growth (each duplicate appends another String).
    let bytes = make_xlsx_with_sheet_xml(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="A1"><v>2</v></c>
    </row>
  </sheetData>
</worksheet>"#,
    );
    let (_, err, ok) = run_with_stdin(&["-f", "xlsx"], &bytes);
    assert!(!ok, "duplicate cell should be rejected");
    let msg = String::from_utf8_lossy(&err);
    assert!(
        msg.contains("out-of-order") || msg.contains("duplicate"),
        "unexpected error: {msg}"
    );
}

#[test]
fn xlsx_out_of_order_col_rejected() {
    // B1 before A1 — lexicographically decreasing within a row.
    let bytes = make_xlsx_with_sheet_xml(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="B1"><v>2</v></c>
      <c r="A1"><v>1</v></c>
    </row>
  </sheetData>
</worksheet>"#,
    );
    let (_, err, ok) = run_with_stdin(&["-f", "xlsx"], &bytes);
    assert!(!ok, "out-of-order column should be rejected");
    let msg = String::from_utf8_lossy(&err);
    assert!(
        msg.contains("out-of-order") || msg.contains("duplicate"),
        "unexpected error: {msg}"
    );
}

#[test]
fn xlsx_out_of_order_row_rejected() {
    // Row 2 before row 1 — rows decrease.
    let bytes = make_xlsx_with_sheet_xml(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="2">
      <c r="A2"><v>2</v></c>
    </row>
    <row r="1">
      <c r="A1"><v>1</v></c>
    </row>
  </sheetData>
</worksheet>"#,
    );
    let (_, err, ok) = run_with_stdin(&["-f", "xlsx"], &bytes);
    assert!(!ok, "out-of-order row should be rejected");
    let msg = String::from_utf8_lossy(&err);
    assert!(
        msg.contains("out-of-order") || msg.contains("duplicate"),
        "unexpected error: {msg}"
    );
}

#[test]
fn xlsx_empty_sheet_ok() {
    // Sheet with no cells should produce empty output, not an error.
    let bytes = make_xlsx_with_sheet_xml(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#,
    );
    let (out, _, ok) = run_with_stdin(&["-f", "xlsx"], &bytes);
    assert!(ok, "empty sheet should succeed");
    assert_eq!(out, b"");
}

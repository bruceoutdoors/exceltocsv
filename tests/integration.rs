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
fn quoting_none_requires_escapechar() {
    let (_, err, ok) = run(&["-U", "3", "tests/fixtures/simple.xlsx"]);
    assert!(!ok);
    assert!(
        String::from_utf8_lossy(&err).contains("--out-escapechar"),
        "error message should mention --out-escapechar"
    );
}

#[test]
fn quoting_none_with_escapechar_ok() {
    let (out, _, ok) = run(&["-U", "3", "-P", "\\", "tests/fixtures/simple.xlsx"]);
    assert!(ok);
    assert!(!out.is_empty());
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

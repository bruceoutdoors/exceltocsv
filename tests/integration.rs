use std::io::Write;
use std::path::Path;
use std::process::{Command, Stdio};

fn bin() -> Command {
    Command::new(env!("CARGO_BIN_EXE_exceltocsv"))
}

fn run(args: &[&str]) -> (String, String, bool) {
    run_in(Path::new("."), args)
}

fn run_in(dir: &Path, args: &[&str]) -> (String, String, bool) {
    let out = bin()
        .args(args)
        .current_dir(dir)
        .output()
        .expect("failed to run exceltocsv");
    (
        String::from_utf8_lossy(&out.stdout).into_owned(),
        String::from_utf8_lossy(&out.stderr).into_owned(),
        out.status.success(),
    )
}

fn run_with_stdin(args: &[&str], stdin_bytes: &[u8]) -> (String, String, bool) {
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
    (
        String::from_utf8_lossy(&out.stdout).into_owned(),
        String::from_utf8_lossy(&out.stderr).into_owned(),
        out.status.success(),
    )
}

fn expected(name: &str) -> String {
    std::fs::read_to_string(Path::new("tests/expected").join(name))
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
    assert_eq!(out.trim(), "Sheet1\nSheet2");
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
fn tab_delimiter() {
    let (out, _, ok) = run(&["--tabs", "tests/fixtures/simple.xlsx"]);
    assert!(ok);
    let expected_tsv = expected("simple.csv").replace(',', "\t");
    assert_eq!(out, expected_tsv);
}

#[test]
fn custom_delimiter() {
    let (out, _, ok) = run(&["-d", "|", "tests/fixtures/simple.xlsx"]);
    assert!(ok);
    let expected_psv = expected("simple.csv").replace(',', "|");
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
    assert!(ok, "{err}");

    let s1 = std::fs::read_to_string(dir.path().join("sheet1.csv")).expect("sheet1.csv missing");
    let s2 = std::fs::read_to_string(dir.path().join("sheet2.csv")).expect("sheet2.csv missing");
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
    assert!(ok, "{err}");

    let s1 = std::fs::read_to_string(dir.path().join("Sheet1.csv")).expect("Sheet1.csv missing");
    let s2 = std::fs::read_to_string(dir.path().join("Sheet2.csv")).expect("Sheet2.csv missing");
    assert_eq!(s1, expected("multiple_sheets_sheet1.csv"));
    assert_eq!(s2, expected("multiple_sheets_sheet2.csv"));
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

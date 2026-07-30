use std::path::Path;
use std::process::Command;

fn bin() -> Command {
    Command::new(env!("CARGO_BIN_EXE_exceltocsv"))
}

fn run(args: &[&str]) -> (String, String, bool) {
    let out = bin().args(args).output().expect("failed to run exceltocsv");
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
    let out = bin()
        .args(["--write-sheets", "all", fixture.to_str().unwrap()])
        .current_dir(dir.path())
        .output()
        .expect("run");
    assert!(
        out.status.success(),
        "{}",
        String::from_utf8_lossy(&out.stderr)
    );
    assert!(
        dir.path().join("sheet1.csv").exists(),
        "sheet1.csv not created"
    );
    assert!(
        dir.path().join("sheet2.csv").exists(),
        "sheet2.csv not created"
    );
}

#[test]
fn write_sheets_use_names() {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/multiple_sheets.xlsx");
    let dir = tempfile::tempdir().expect("tempdir");
    let out = bin()
        .args([
            "--write-sheets",
            "all",
            "--use-sheet-names",
            fixture.to_str().unwrap(),
        ])
        .current_dir(dir.path())
        .output()
        .expect("run");
    assert!(
        out.status.success(),
        "{}",
        String::from_utf8_lossy(&out.stderr)
    );
    assert!(
        dir.path().join("Sheet1.csv").exists(),
        "Sheet1.csv not created"
    );
    assert!(
        dir.path().join("Sheet2.csv").exists(),
        "Sheet2.csv not created"
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

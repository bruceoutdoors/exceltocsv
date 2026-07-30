use std::path::Path;

use anyhow::Result;
use rust_xlsxwriter::{ExcelDateTime, Format, Workbook};

fn main() -> Result<()> {
    std::fs::create_dir_all("tests/fixtures")?;

    simple_xlsx()?;
    multiple_sheets_xlsx()?;
    quoted_values_xlsx()?;
    unicode_xlsx()?;
    types_xlsx()?;

    println!("XLSX fixtures written to tests/fixtures/");
    Ok(())
}

fn simple_xlsx() -> Result<()> {
    let mut wb = Workbook::new();
    let ws = wb.add_worksheet();
    ws.write(0, 0, "Name")?;
    ws.write(0, 1, "Amount")?;
    ws.write(0, 2, "Active")?;
    ws.write(1, 0, "Alice")?;
    ws.write(1, 1, 100_i32)?;
    ws.write(1, 2, "true")?;
    ws.write(2, 0, "Bob")?;
    ws.write(2, 1, 250.5_f64)?;
    ws.write(2, 2, "false")?;
    wb.save(Path::new("tests/fixtures/simple.xlsx"))?;
    Ok(())
}

fn multiple_sheets_xlsx() -> Result<()> {
    let mut wb = Workbook::new();

    let ws1 = wb.add_worksheet();
    ws1.set_name("Sheet1")?;
    ws1.write(0, 0, "ID")?;
    ws1.write(0, 1, "Value")?;
    ws1.write(1, 0, 1_i32)?;
    ws1.write(1, 1, "Alpha")?;
    ws1.write(2, 0, 2_i32)?;
    ws1.write(2, 1, "Beta")?;

    let ws2 = wb.add_worksheet();
    ws2.set_name("Sheet2")?;
    ws2.write(0, 0, "Color")?;
    ws2.write(0, 1, "Hex")?;
    ws2.write(1, 0, "Red")?;
    ws2.write(1, 1, "#FF0000")?;
    ws2.write(2, 0, "Green")?;
    ws2.write(2, 1, "#00FF00")?;

    wb.save(Path::new("tests/fixtures/multiple_sheets.xlsx"))?;
    Ok(())
}

fn quoted_values_xlsx() -> Result<()> {
    let mut wb = Workbook::new();
    let ws = wb.add_worksheet();
    ws.write(0, 0, "Name")?;
    ws.write(0, 1, "Note")?;
    ws.write(1, 0, "Alice")?;
    ws.write(1, 1, "hello, world")?;
    ws.write(2, 0, "Bob")?;
    ws.write(2, 1, "line one\nline two")?;
    ws.write(3, 0, "Charlie")?;
    ws.write(3, 1, "he said \"hi\"")?;
    wb.save(Path::new("tests/fixtures/quoted_values.xlsx"))?;
    Ok(())
}

fn unicode_xlsx() -> Result<()> {
    let mut wb = Workbook::new();
    let ws = wb.add_worksheet();
    ws.write(0, 0, "Name")?;
    ws.write(0, 1, "City")?;
    ws.write(1, 0, "Ali")?;
    ws.write(1, 1, "Kuala Lumpur")?;
    ws.write(2, 0, "Mei Ling")?;
    ws.write(2, 1, "槟城")?;
    ws.write(3, 0, "Siti")?;
    ws.write(3, 1, "Johor Bahru")?;
    wb.save(Path::new("tests/fixtures/unicode.xlsx"))?;
    Ok(())
}

fn types_xlsx() -> Result<()> {
    let mut wb = Workbook::new();
    let ws = wb.add_worksheet();
    ws.set_name("Types")?;

    ws.write(0, 0, "type")?;
    ws.write(0, 1, "value")?;

    // Native boolean cells
    ws.write(1, 0, "bool_true")?;
    ws.write(1, 1, true)?;
    ws.write(2, 0, "bool_false")?;
    ws.write(2, 1, false)?;

    // Empty cell (sparse — col 1 intentionally omitted)
    ws.write(3, 0, "empty")?;

    // Integer and float
    ws.write(4, 0, "integer")?;
    ws.write(4, 1, 42_i32)?;
    ws.write(5, 0, "float")?;
    ws.write(5, 1, 3.14_f64)?;

    // Date cell with format so calamine reads it as DateTime
    let date_fmt = Format::new().set_num_format("yyyy-mm-dd");
    let date = ExcelDateTime::from_ymd(2024, 1, 15)?;
    ws.write(6, 0, "date")?;
    ws.write_with_format(6, 1, &date, &date_fmt)?;

    // Datetime with milliseconds
    let dt_fmt = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss.000");
    let dt_with_ms = ExcelDateTime::from_ymd(2024, 3, 15)?.and_hms_milli(14, 30, 45, 500)?;
    ws.write(7, 0, "datetime_ms")?;
    ws.write_with_format(7, 1, &dt_with_ms, &dt_fmt)?;

    // Leading-zero string (written as text, not number)
    ws.write(8, 0, "leading_zero")?;
    ws.write(8, 1, "007")?;

    wb.save(Path::new("tests/fixtures/types.xlsx"))?;
    Ok(())
}

use std::fs::File;
use std::io::{self, BufWriter, Write};
use std::path::PathBuf;

use anyhow::{bail, Context, Result};
use calamine::{open_workbook_auto, Data, Range, Reader};
use clap::Parser;
use csv::WriterBuilder;

#[derive(Parser, Debug)]
#[command(name = "exceltocsv", about = "Convert Excel files to CSV", version)]
struct Args {
    /// Input Excel file (.xls, .xlsx, .xlsb, .ods)
    input: PathBuf,

    /// Explicit format override: xls or xlsx
    #[arg(short = 'f', long, value_name = "FMT")]
    format: Option<String>,

    /// Print worksheet names to stdout and exit
    #[arg(short = 'n', long)]
    names: bool,

    /// Select worksheet by name (default: first sheet)
    #[arg(long, value_name = "NAME")]
    sheet: Option<String>,

    /// Comma-separated sheet names (or "all") to write to individual .csv files
    #[arg(long = "write-sheets", value_name = "SHEETS")]
    write_sheets: Option<String>,

    /// Use sheet names as output filenames (used with --write-sheets)
    #[arg(long = "use-sheet-names")]
    use_sheet_names: bool,

    /// Ignored in v0.1.0: calamine always scans declared dimensions
    #[arg(long = "reset-dimensions")]
    reset_dimensions: bool,

    /// Encoding hint for legacy XLS files (informational in v0.1.0; calamine defaults to CP1252)
    #[arg(long = "encoding-xls", value_name = "ENCODING")]
    encoding_xls: Option<String>,

    /// Output CSV delimiter character (default: comma)
    #[arg(short = 'd', long, value_name = "CHAR")]
    delimiter: Option<char>,

    /// Use tab as delimiter (shorthand for -d '\\t')
    #[arg(short = 't', long)]
    tabs: bool,

    /// CSV quote character (default: double-quote)
    #[arg(short = 'q', long, value_name = "CHAR")]
    quotechar: Option<char>,

    /// Quoting mode: minimal | all | nonnumeric | none
    #[arg(short = 'u', long, value_name = "MODE")]
    quoting: Option<String>,

    /// Disable double-quote escaping; use escape character instead
    #[arg(short = 'b', long = "no-doublequote")]
    no_doublequote: bool,

    /// Escape character (used when --no-doublequote is set)
    #[arg(short = 'p', long, value_name = "CHAR")]
    escapechar: Option<char>,

    /// Field size limit in bytes (informational)
    #[arg(short = 'z', long = "field-size-limit", value_name = "N")]
    field_size_limit: Option<u64>,
}

fn main() -> Result<()> {
    let args = Args::parse();

    if args.reset_dimensions {
        eprintln!(
            "warning: --reset-dimensions is not yet implemented; \
             calamine always uses declared worksheet dimensions"
        );
    }
    if args.encoding_xls.is_some() {
        eprintln!(
            "warning: --encoding-xls is informational in v0.1.0; \
             calamine uses CP1252 as the XLS fallback encoding"
        );
    }
    if args.no_doublequote && args.escapechar.is_none() {
        bail!("--no-doublequote requires --escapechar to be set");
    }

    let mut workbook = open_workbook_auto(&args.input)
        .with_context(|| format!("cannot open '{}'", args.input.display()))?;

    let sheet_names = workbook.sheet_names().to_vec();

    if args.names {
        for name in &sheet_names {
            println!("{name}");
        }
        return Ok(());
    }

    if let Some(ref spec) = args.write_sheets {
        let targets: Vec<&str> = if spec.eq_ignore_ascii_case("all") {
            sheet_names.iter().map(String::as_str).collect()
        } else {
            spec.split(',').map(str::trim).collect()
        };

        for name in targets {
            let range = workbook
                .worksheet_range(name)
                .with_context(|| format!("sheet '{name}' not found"))?;

            let filename = if args.use_sheet_names {
                format!("{name}.csv")
            } else {
                let idx = sheet_names.iter().position(|s| s == name).unwrap_or(0);
                format!("sheet{}.csv", idx + 1)
            };

            let file =
                File::create(&filename).with_context(|| format!("cannot create '{filename}'"))?;
            write_range(&range, BufWriter::new(file), &args)?;
        }
        return Ok(());
    }

    let sheet_name = args
        .sheet
        .as_deref()
        .or_else(|| sheet_names.first().map(String::as_str))
        .context("workbook has no sheets")?;

    let range = workbook
        .worksheet_range(sheet_name)
        .with_context(|| format!("sheet '{sheet_name}' not found"))?;

    let stdout = io::stdout();
    write_range(&range, stdout.lock(), &args)
}

fn write_range<W: Write>(range: &Range<Data>, writer: W, args: &Args) -> Result<()> {
    let delimiter = resolve_delimiter(args)?;
    let quote = resolve_quote(args)?;
    let quoting = resolve_quoting(args);

    let mut builder = WriterBuilder::new();
    builder
        .delimiter(delimiter)
        .quote(quote)
        .quote_style(quoting)
        .double_quote(!args.no_doublequote)
        .terminator(csv::Terminator::Any(b'\n'));

    if let Some(esc) = args.escapechar {
        builder.escape(ascii_byte(esc, "--escapechar")?);
    }

    let mut wtr = builder.from_writer(writer);
    for row in range.rows() {
        wtr.write_record(row.iter().map(render_cell))?;
    }
    wtr.flush()?;
    Ok(())
}

fn resolve_delimiter(args: &Args) -> Result<u8> {
    if args.tabs {
        return Ok(b'\t');
    }
    match args.delimiter {
        Some(c) => ascii_byte(c, "--delimiter"),
        None => Ok(b','),
    }
}

fn resolve_quote(args: &Args) -> Result<u8> {
    match args.quotechar {
        Some(c) => ascii_byte(c, "--quotechar"),
        None => Ok(b'"'),
    }
}

fn resolve_quoting(args: &Args) -> csv::QuoteStyle {
    match args.quoting.as_deref() {
        Some("all") => csv::QuoteStyle::Always,
        Some("nonnumeric") => csv::QuoteStyle::NonNumeric,
        Some("none") => csv::QuoteStyle::Never,
        _ => csv::QuoteStyle::Necessary,
    }
}

fn ascii_byte(c: char, flag: &str) -> Result<u8> {
    if c.is_ascii() {
        Ok(c as u8)
    } else {
        bail!("{flag} must be an ASCII character, got '{c}'")
    }
}

fn render_cell(cell: &Data) -> String {
    match cell {
        Data::Empty => String::new(),
        Data::String(s) | Data::DateTimeIso(s) | Data::DurationIso(s) => s.clone(),
        Data::Int(i) => i.to_string(),
        Data::Float(f) => format!("{f}"),
        Data::DateTime(d) => format!("{}", d.as_f64()),
        Data::Bool(b) => if *b { "true" } else { "false" }.to_string(),
        Data::Error(_) => "#ERROR".to_string(),
    }
}

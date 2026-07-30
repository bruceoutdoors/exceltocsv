use std::fs::File;
use std::io::{self, BufWriter, Write};
use std::path::PathBuf;

use anyhow::{Context, Result};
use calamine::{open_workbook_auto, Data, Range, Reader};
use clap::{Parser, ValueEnum};
use csv::WriterBuilder;

#[derive(Parser, Debug)]
#[command(name = "exceltocsv", about = "Convert Excel files to CSV", version)]
struct Args {
    /// Input Excel file (.xls, .xlsx, .xlsb, .ods)
    input: PathBuf,

    /// Explicit format override (informational in v0.1.0; format is auto-detected)
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

    /// Ignored in v0.1.0: calamine always uses declared worksheet dimensions
    #[arg(long = "reset-dimensions")]
    reset_dimensions: bool,

    /// Encoding hint for legacy XLS files (informational in v0.1.0; calamine defaults to CP1252)
    #[arg(long = "encoding-xls", value_name = "ENCODING")]
    encoding_xls: Option<String>,

    /// Output CSV delimiter character (default: comma)
    #[arg(short = 'd', long, value_name = "CHAR", value_parser = parse_ascii_byte)]
    delimiter: Option<u8>,

    /// Use tab as delimiter (shorthand for -d '\t')
    #[arg(short = 't', long, conflicts_with = "delimiter")]
    tabs: bool,

    /// CSV quote character (default: double-quote)
    #[arg(short = 'q', long, value_name = "CHAR", value_parser = parse_ascii_byte)]
    quotechar: Option<u8>,

    /// Quoting mode
    #[arg(short = 'u', long, value_name = "MODE")]
    quoting: Option<QuoteMode>,

    /// Disable double-quote escaping; use escape character instead
    #[arg(short = 'b', long = "no-doublequote", requires = "escapechar")]
    no_doublequote: bool,

    /// Escape character (used when --no-doublequote is set)
    #[arg(short = 'p', long, value_name = "CHAR", value_parser = parse_ascii_byte)]
    escapechar: Option<u8>,

    /// Field size limit in bytes (informational; no enforcement in v0.1.0)
    #[arg(short = 'z', long = "field-size-limit", value_name = "N")]
    field_size_limit: Option<u64>,
}

#[derive(Clone, Debug, ValueEnum)]
enum QuoteMode {
    Minimal,
    All,
    Nonnumeric,
    None,
}

fn main() -> Result<()> {
    let args = Args::parse();

    if args.format.is_some() {
        eprintln!(
            "warning: --format is informational in v0.1.0; \
             format is detected automatically from the file extension"
        );
    }
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
    if args.field_size_limit.is_some() {
        eprintln!("warning: --field-size-limit is informational in v0.1.0; no limit is enforced");
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
    let delimiter = if args.tabs {
        b'\t'
    } else {
        args.delimiter.unwrap_or(b',')
    };
    let quote = args.quotechar.unwrap_or(b'"');
    let quote_style = match &args.quoting {
        Some(QuoteMode::All) => csv::QuoteStyle::Always,
        Some(QuoteMode::Nonnumeric) => csv::QuoteStyle::NonNumeric,
        Some(QuoteMode::None) => csv::QuoteStyle::Never,
        _ => csv::QuoteStyle::Necessary,
    };

    let mut builder = WriterBuilder::new();
    builder
        .delimiter(delimiter)
        .quote(quote)
        .quote_style(quote_style)
        .double_quote(!args.no_doublequote)
        .terminator(csv::Terminator::Any(b'\n'));

    if let Some(esc) = args.escapechar {
        builder.escape(esc);
    }

    let mut wtr = builder.from_writer(writer);
    for row in range.rows() {
        wtr.write_record(row.iter().map(render_cell))?;
    }
    wtr.flush()?;
    Ok(())
}

fn parse_ascii_byte(s: &str) -> Result<u8, String> {
    let mut chars = s.chars();
    match (chars.next(), chars.next()) {
        (Some(c), None) if c.is_ascii() => Ok(c as u8),
        (Some(c), None) => Err(format!("'{c}' is not an ASCII character")),
        _ => Err(format!("{s:?} must be a single ASCII character")),
    }
}

fn render_cell(cell: &Data) -> String {
    match cell {
        Data::Empty => String::new(),
        Data::String(s) | Data::DateTimeIso(s) | Data::DurationIso(s) => s.clone(),
        Data::Int(i) => i.to_string(),
        Data::Float(f) => f.to_string(),
        Data::DateTime(d) => d.as_f64().to_string(),
        Data::Bool(b) => b.to_string(),
        Data::Error(_) => "#ERROR".to_string(),
    }
}

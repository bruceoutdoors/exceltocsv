use std::collections::HashMap;
use std::io::{self, BufWriter, Cursor, Read, Write};
use std::path::{Path, PathBuf};

use anyhow::{Context, Result};
use calamine::{
    open_workbook_auto_from_rs, open_workbook_from_rs, Data, Range, Reader, Sheets, Xls, Xlsx,
};
use clap::{Parser, ValueEnum};
use csv::WriterBuilder;

#[derive(Parser, Debug)]
#[command(name = "exceltocsv", about = "Convert Excel files to CSV", version)]
struct Args {
    /// Input Excel file (.xls, .xlsx); omit or use - to read from stdin
    input: Option<PathBuf>,

    /// Force input format: xls or xlsx (required for stdin if auto-detection fails)
    #[arg(short = 'f', long, value_name = "FMT")]
    format: Option<Format>,

    /// Print worksheet names to stdout and exit
    #[arg(short = 'n', long)]
    names: bool,

    /// Select worksheet by name (default: first sheet)
    #[arg(long, value_name = "NAME")]
    sheet: Option<String>,

    /// Write all sheets to .csv files; use - for all, or comma-separated sheet names
    #[arg(long = "write-sheets", value_name = "SHEETS")]
    write_sheets: Option<String>,

    /// Use sheet names as output filenames (requires --write-sheets)
    #[arg(long = "use-sheet-names", requires = "write_sheets")]
    use_sheet_names: bool,

    /// Output CSV delimiter character (default: comma)
    #[arg(short = 'd', long, value_name = "CHAR", value_parser = parse_ascii_byte)]
    delimiter: Option<u8>,

    /// Use tab as delimiter
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
}

#[derive(Clone, Debug, ValueEnum)]
enum Format {
    Xls,
    Xlsx,
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
    let bytes = load_input(&args.input)?;
    let cursor = Cursor::new(bytes);

    let mut wb: Sheets<Cursor<Vec<u8>>> = match args.format {
        Some(Format::Xls) => open_workbook_from_rs::<Xls<_>, _>(cursor)
            .map(Sheets::Xls)
            .map_err(|e| anyhow::anyhow!("failed to open workbook as XLS: {e}"))?,
        Some(Format::Xlsx) => open_workbook_from_rs::<Xlsx<_>, _>(cursor)
            .map(Sheets::Xlsx)
            .map_err(|e| anyhow::anyhow!("failed to open workbook as XLSX: {e}"))?,
        None => open_workbook_auto_from_rs(cursor).map_err(|_| {
            anyhow::anyhow!(
                "cannot detect workbook format; try specifying --format xls or --format xlsx"
            )
        })?,
    };

    process_workbook(&mut wb, &args)
}

fn load_input(input: &Option<PathBuf>) -> Result<Vec<u8>> {
    match input {
        None => read_stdin(),
        Some(p) if p == Path::new("-") => read_stdin(),
        Some(p) => std::fs::read(p).with_context(|| format!("cannot open '{}'", p.display())),
    }
}

fn read_stdin() -> Result<Vec<u8>> {
    let mut buf = Vec::new();
    io::stdin()
        .read_to_end(&mut buf)
        .context("failed to read stdin")?;
    Ok(buf)
}

fn process_workbook(wb: &mut Sheets<Cursor<Vec<u8>>>, args: &Args) -> Result<()> {
    let sheet_names = wb.sheet_names().to_vec();

    if args.names {
        for name in &sheet_names {
            println!("{name}");
        }
        return Ok(());
    }

    if let Some(ref spec) = args.write_sheets {
        let targets: Vec<String> = if spec == "-" {
            sheet_names.clone()
        } else {
            spec.split(',').map(|s| s.trim().to_owned()).collect()
        };

        let mut used: HashMap<String, usize> = HashMap::new();
        for name in &targets {
            let range = wb
                .worksheet_range(name)
                .with_context(|| format!("sheet '{name}' not found"))?;

            let base = if args.use_sheet_names {
                sanitize_sheet_name(name)
            } else {
                let idx = sheet_names.iter().position(|s| s == name).unwrap_or(0);
                format!("sheet{}", idx + 1)
            };

            let count = used.entry(base.clone()).or_insert(0);
            let filename = if *count == 0 {
                format!("{base}.csv")
            } else {
                format!("{base}_{count}.csv")
            };
            *count += 1;

            let file = std::fs::File::create(&filename)
                .with_context(|| format!("cannot create '{filename}'"))?;
            write_range(&range, BufWriter::new(file), args)?;
        }
        return Ok(());
    }

    let sheet_name = args
        .sheet
        .as_deref()
        .or_else(|| sheet_names.first().map(String::as_str))
        .context("workbook has no sheets")?;

    let range = wb
        .worksheet_range(sheet_name)
        .with_context(|| format!("sheet '{sheet_name}' not found"))?;

    let stdout = io::stdout();
    write_range(&range, stdout.lock(), args)
}

fn sanitize_sheet_name(name: &str) -> String {
    let s: String = name
        .chars()
        .map(|c| match c {
            '/' | '\\' | ':' | '*' | '?' | '"' | '<' | '>' | '|' => '_',
            c if c.is_control() => '_',
            c => c,
        })
        .collect();

    let s = s.trim();
    match s {
        "" | "." | ".." => "sheet".to_string(),
        s if s.starts_with('.') => format!("_{}", &s[1..]),
        s => s.to_string(),
    }
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

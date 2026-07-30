use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::{self, BufReader, Cursor, Read, Seek, Write};
use std::path::{Path, PathBuf};

use anyhow::{Context, Result};
use calamine::{
    open_workbook_auto, open_workbook_auto_from_rs, open_workbook_from_rs, Data, Range, Reader,
    Sheets, Xls, Xlsx,
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
    #[arg(
        short = 'n',
        long,
        conflicts_with = "sheet",
        conflicts_with = "write_sheets"
    )]
    names: bool,

    /// Select worksheet by name (default: first sheet)
    #[arg(long, value_name = "NAME", conflicts_with = "write_sheets")]
    sheet: Option<String>,

    /// Write sheets to .csv files; - for all, or comma-separated names
    #[arg(long = "write-sheets", value_name = "SHEETS")]
    write_sheets: Option<String>,

    /// Use sheet names as output filenames (requires --write-sheets)
    #[arg(long = "use-sheet-names", requires = "write_sheets")]
    use_sheet_names: bool,

    /// Output CSV delimiter character (default: comma)
    #[arg(short = 'D', long = "out-delimiter", value_name = "CHAR", value_parser = parse_ascii_byte)]
    out_delimiter: Option<u8>,

    /// Use tab as delimiter
    #[arg(short = 'T', long = "out-tabs", conflicts_with = "out_delimiter")]
    out_tabs: bool,

    /// CSV quote character (default: double-quote)
    #[arg(short = 'Q', long = "out-quotechar", value_name = "CHAR", value_parser = parse_ascii_byte)]
    out_quotechar: Option<u8>,

    /// Quoting mode: 0=minimal 1=all 2=nonnumeric 3=none (mode 3 requires --out-escapechar)
    #[arg(short = 'U', long = "out-quoting", value_name = "MODE")]
    out_quoting: Option<u8>,

    /// Disable double-quote escaping; use escape character instead
    #[arg(short = 'B', long = "out-no-doublequote", requires = "out_escapechar")]
    out_no_doublequote: bool,

    /// Escape character (used with --out-no-doublequote or --out-quoting 3)
    #[arg(short = 'P', long = "out-escapechar", value_name = "CHAR", value_parser = parse_ascii_byte)]
    out_escapechar: Option<u8>,

    /// Line terminator: lf (default) or crlf
    #[arg(short = 'M', long = "out-lineterminator", value_name = "EOL")]
    out_lineterminator: Option<LineTerminator>,
}

#[derive(Clone, Debug, ValueEnum)]
enum Format {
    Xls,
    Xlsx,
}

#[derive(Clone, Debug, ValueEnum)]
enum LineTerminator {
    Lf,
    Crlf,
}

fn main() -> Result<()> {
    let args = Args::parse();

    if args.out_quoting == Some(3) && args.out_escapechar.is_none() {
        anyhow::bail!("--out-quoting 3 (none) requires --out-escapechar");
    }
    if let Some(q) = args.out_quoting {
        if q > 3 {
            anyhow::bail!("--out-quoting must be 0-3 (0=minimal, 1=all, 2=nonnumeric, 3=none)");
        }
    }

    let is_stdin = args.input.is_none() || args.input.as_deref() == Some(Path::new("-"));

    if is_stdin {
        let mut buf = Vec::new();
        io::stdin()
            .read_to_end(&mut buf)
            .context("failed to read stdin")?;
        let cursor = Cursor::new(buf);
        let mut wb: Sheets<Cursor<Vec<u8>>> = match &args.format {
            Some(Format::Xls) => open_workbook_from_rs::<Xls<_>, _>(cursor)
                .map(Sheets::Xls)
                .map_err(|e| anyhow::anyhow!("failed to open as XLS: {e}"))?,
            Some(Format::Xlsx) => open_workbook_from_rs::<Xlsx<_>, _>(cursor)
                .map(Sheets::Xlsx)
                .map_err(|e| anyhow::anyhow!("failed to open as XLSX: {e}"))?,
            None => open_workbook_auto_from_rs(cursor)
                .map_err(|_| anyhow::anyhow!("cannot detect workbook format; try --format"))?,
        };
        process_workbook(&mut wb, &args)
    } else {
        let path = args.input.as_deref().unwrap();
        let mut wb: Sheets<BufReader<File>> = match &args.format {
            Some(fmt) => {
                let f = File::open(path)
                    .with_context(|| format!("cannot open '{}'", path.display()))?;
                let br = BufReader::new(f);
                match fmt {
                    Format::Xls => open_workbook_from_rs::<Xls<_>, _>(br)
                        .map(Sheets::Xls)
                        .map_err(|e| anyhow::anyhow!("failed to open as XLS: {e}"))?,
                    Format::Xlsx => open_workbook_from_rs::<Xlsx<_>, _>(br)
                        .map(Sheets::Xlsx)
                        .map_err(|e| anyhow::anyhow!("failed to open as XLSX: {e}"))?,
                }
            }
            None => open_workbook_auto(path)
                .with_context(|| format!("cannot open '{}'", path.display()))?,
        };
        process_workbook(&mut wb, &args)
    }
}

fn process_workbook<RS: Read + Seek>(wb: &mut Sheets<RS>, args: &Args) -> Result<()> {
    let sheet_names = wb.sheet_names();

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

        // Pre-allocate all filenames upfront for complete collision detection
        let mut filename_set: HashSet<String> = HashSet::new();
        let mut plan: Vec<(String, String)> = Vec::with_capacity(targets.len());
        let mut base_counts: HashMap<String, usize> = HashMap::new();

        for name in &targets {
            let base = if args.use_sheet_names {
                sanitize_sheet_name(name)
            } else {
                let idx = sheet_names.iter().position(|s| s == name).unwrap_or(0);
                format!("sheet{}", idx + 1)
            };

            let counter = base_counts.entry(base.clone()).or_insert(0);
            let filename = loop {
                let candidate = if *counter == 0 {
                    format!("{base}.csv")
                } else {
                    format!("{base}_{counter}.csv")
                };
                *counter += 1;
                if !filename_set.contains(&candidate) {
                    break candidate;
                }
            };
            filename_set.insert(filename.clone());
            plan.push((name.clone(), filename));
        }

        // Write files, cleaning up any created files on failure
        let mut created: Vec<String> = Vec::new();
        let result = (|| -> Result<()> {
            for (name, filename) in &plan {
                let range = wb
                    .worksheet_range(name)
                    .with_context(|| format!("sheet '{name}' not found"))?;

                let file = File::create_new(filename)
                    .with_context(|| format!("cannot create '{filename}': file already exists"))?;
                created.push(filename.clone());

                write_range(&range, file, args)?;
            }
            Ok(())
        })();

        if let Err(e) = result {
            for f in &created {
                let _ = std::fs::remove_file(f);
            }
            return Err(e);
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
    let delimiter = if args.out_tabs {
        b'\t'
    } else {
        args.out_delimiter.unwrap_or(b',')
    };
    let quote = args.out_quotechar.unwrap_or(b'"');
    let quote_style = match args.out_quoting {
        Some(1) => csv::QuoteStyle::Always,
        Some(2) => csv::QuoteStyle::NonNumeric,
        Some(3) => csv::QuoteStyle::Never,
        _ => csv::QuoteStyle::Necessary,
    };
    let terminator = match &args.out_lineterminator {
        Some(LineTerminator::Crlf) => csv::Terminator::CRLF,
        _ => csv::Terminator::Any(b'\n'),
    };

    let mut builder = WriterBuilder::new();
    builder
        .delimiter(delimiter)
        .quote(quote)
        .quote_style(quote_style)
        .double_quote(!args.out_no_doublequote)
        .terminator(terminator);

    if let Some(esc) = args.out_escapechar {
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
        Data::Bool(b) => if *b { "true" } else { "false" }.to_string(),
        Data::Error(e) => e.to_string(),
        Data::DateTime(d) => render_datetime(d),
    }
}

fn render_datetime(d: &calamine::ExcelDateTime) -> String {
    if d.is_duration() {
        let total_secs = (d.as_f64() * 86400.0).round() as i64;
        let h = total_secs / 3600;
        let m = (total_secs % 3600) / 60;
        let s = total_secs % 60;
        format!("{h:02}:{m:02}:{s:02}")
    } else {
        let (year, month, day, hour, min, sec, _) = d.to_ymd_hms_milli();
        if hour == 0 && min == 0 && sec == 0 {
            format!("{year:04}-{month:02}-{day:02}")
        } else {
            format!("{year:04}-{month:02}-{day:02}T{hour:02}:{min:02}:{sec:02}")
        }
    }
}

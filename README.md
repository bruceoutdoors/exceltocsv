# exceltocsv

Converts Excel files (XLS, XLSX) to CSV output.

## Why this exists

User-uploaded Excel files need to be converted to CSV inside a narrow, embeddable execution boundary. `exceltocsv` runs as a WASI module so it can be embedded in a host application with no OS-level process spawning. It also builds as a native binary for local testing.

## Why it is intentionally small

This is not a general-purpose spreadsheet engine. It is a small, auditable Excel-to-CSV converter. The entire conversion path is: read Excel with `calamine`, stream rows through the `csv` crate, write to stdout. Nothing more.

## Install

Download the latest binary from the [GitHub Releases page](../../releases).

**Linux (native binary):**

```bash
curl -L https://github.com/bruceoutdoors/exceltocsv/releases/latest/download/exceltocsv-linux-amd64 -o exceltocsv
chmod +x exceltocsv
```

Verify the checksum:

```bash
curl -L https://github.com/bruceoutdoors/exceltocsv/releases/latest/download/exceltocsv-linux-amd64.sha256 | sha256sum -c
```

**WASI module:**

Download `exceltocsv.wasm` and `exceltocsv.wasm.sha256` from the releases page.

## Usage

```bash
# Convert first sheet to CSV on stdout
exceltocsv input.xlsx > output.csv

# List all worksheet names
exceltocsv --names input.xlsx

# Select a specific sheet
exceltocsv --sheet "Sales Q1" input.xlsx > sales.csv

# Tab-delimited output
exceltocsv --tabs input.xlsx > output.tsv

# Custom delimiter
exceltocsv -d '|' input.xlsx > output.psv

# Write all sheets to individual CSV files using sheet names
exceltocsv --write-sheets all --use-sheet-names input.xlsx
```

Full option reference:

```text
-f, --format <FMT>         Explicit format: xls or xlsx
-n, --names                List worksheet names and exit
    --sheet <NAME>         Select worksheet by name (default: first sheet)
    --write-sheets <S>     Comma-separated sheet names or "all"; write each to a file
    --use-sheet-names      Use sheet names as output filenames (with --write-sheets)
    --reset-dimensions     Accepted but not yet implemented (calamine limitation)
    --encoding-xls <E>     Informational in v0.1.0; calamine defaults to CP1252 for XLS
-d, --delimiter <CHAR>     Output delimiter (default: comma)
-t, --tabs                 Use tab as delimiter
-q, --quotechar <CHAR>     Quote character (default: double-quote)
-u, --quoting <MODE>       minimal | all | nonnumeric | none
-b, --no-doublequote       Disable double-quote escaping; requires --escapechar
-p, --escapechar <CHAR>    Escape character (used with --no-doublequote)
-z, --field-size-limit <N> Field size limit in bytes (informational)
```

Note: CSV output uses LF line endings (`\n`). RFC 4180 specifies CRLF but LF is conventional on Unix and avoids test friction.

## Build

**Prerequisites:** Rust >= 1.88, `wasm32-wasip1` target, and `wasm-tools`.

```bash
rustup target add wasm32-wasip1
cargo install wasm-tools
```

**Native binary:**

```bash
cargo build --release
# artifact: target/release/exceltocsv
```

**WASI module:**

```bash
cargo build --target wasm32-wasip1 --release
# artifact: target/wasm32-wasip1/release/exceltocsv.wasm
```

Strip debug symbols to reduce WASM size:

```bash
wasm-tools strip target/wasm32-wasip1/release/exceltocsv.wasm -o exceltocsv.wasm
```

## Testing

Correctness is verified against known XLS/XLSX fixtures with exact expected CSV outputs.

**Fixture files:**

| Fixture | Tests |
| --- | --- |
| `tests/fixtures/simple.xlsx` | Basic XLSX conversion, integer and float rendering |
| `tests/fixtures/simple.xls` | Legacy XLS format parity |
| `tests/fixtures/multiple_sheets.xlsx` | Sheet listing, sheet selection |
| `tests/fixtures/quoted_values.xlsx` | Cells with commas, embedded newlines, double-quotes |
| `tests/fixtures/unicode.xlsx` | CJK and non-ASCII characters |

**Generate XLSX fixtures** (once, after a clean clone):

```bash
cargo run --bin generate_fixtures --features generate-fixtures
```

**Generate the XLS fixture** (requires [uv](https://docs.astral.sh/uv/)):

```bash
uv run tools/generate_fixtures.py
```

**Run tests:**

```bash
cargo test
```

Tests assert exact byte-for-byte CSV output against files in `tests/expected/`.

## Alternatives

If you need a general-purpose Excel-to-CSV converter without the WASI constraint, these tools cover more ground:

- [in2csv](https://csvkit.readthedocs.io/en/latest/scripts/in2csv.html) (Python, part of csvkit) — the CLI that inspired this tool's flag set
- [xsv](https://github.com/BurntSushi/xsv) — fast CSV toolkit in Rust, handles CSV manipulation once you have the CSV
- [ssconvert](https://wiki.gnome.org/Projects/Gnumeric/ssconvert) — part of Gnumeric, converts between many spreadsheet formats

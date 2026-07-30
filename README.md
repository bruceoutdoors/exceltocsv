# exceltocsv

CLI to convert Excel files (XLS, XLSX) to CSV. Available as a Linux binary and as a WASI module for embedding in host runtimes.

## Why

Several capable Excel-to-CSV converters already exist. Lightweight options commonly require a language runtime, while office-suite converters are much larger dependencies, and browser-oriented WebAssembly packages are not standalone WASI command modules. exceltocsv packages this narrow conversion as a Linux CLI and WASI module for applications that want to run it within their own resource and capability limits.

## Install

### Linux

```sh
curl -fL https://github.com/bruceoutdoors/exceltocsv/releases/latest/download/exceltocsv-linux-amd64 \
  -o exceltocsv-linux-amd64
curl -fL https://github.com/bruceoutdoors/exceltocsv/releases/latest/download/exceltocsv-linux-amd64.sha256 \
  | sha256sum -c
chmod +x exceltocsv-linux-amd64
```

### WASI module

```sh
curl -fL https://github.com/bruceoutdoors/exceltocsv/releases/latest/download/exceltocsv.wasm \
  -o exceltocsv.wasm
curl -fL https://github.com/bruceoutdoors/exceltocsv/releases/latest/download/exceltocsv.wasm.sha256 \
  | sha256sum -c
```

## Usage

```
exceltocsv --help
Convert Excel files to CSV

Usage: exceltocsv [OPTIONS] [INPUT]

Arguments:
  [INPUT]  Input Excel file (.xls, .xlsx); omit or use - to read from stdin

Options:
  -f, --format <FMT>              Force input format: xls or xlsx (required for stdin if auto-detection fails) [possible values: xls, xlsx]
  -n, --names                     Print worksheet names to stdout and exit
      --sheet <NAME>              Select worksheet by name (default: first sheet)
      --write-sheets <SHEETS>     Write sheets to .csv files; - for all, or comma-separated names
      --use-sheet-names           Use sheet names as output filenames (requires --write-sheets)
  -D, --out-delimiter <CHAR>      Output CSV delimiter character (default: comma)
  -T, --out-tabs                  Use tab as delimiter
  -Q, --out-quotechar <CHAR>      CSV quote character (default: double-quote)
  -U, --out-quoting <MODE>        Quoting mode: 0=minimal 1=all 2=nonnumeric 3=none (mode 3 requires --out-escapechar)
  -B, --out-no-doublequote        Disable double-quote escaping; use escape character instead
  -P, --out-escapechar <CHAR>     Escape character (used with --out-no-doublequote or --out-quoting 3)
  -M, --out-lineterminator <EOL>  Line terminator: lf (default) or crlf [possible values: lf, crlf]
  -h, --help                      Print help
  -V, --version                   Print version
```

### Examples

```bash
# Convert first sheet to CSV on stdout
exceltocsv input.xlsx > output.csv

# Read from stdin (format auto-detected from magic bytes)
cat input.xlsx | exceltocsv > output.csv
cat input.xls  | exceltocsv -f xls > output.csv

# List all worksheet names
exceltocsv --names input.xlsx

# Select a specific sheet
exceltocsv --sheet "Sales Q1" input.xlsx > sales.csv

# Tab-delimited output
exceltocsv --out-tabs input.xlsx > output.tsv

# Write all sheets to individual CSV files
exceltocsv --write-sheets - --use-sheet-names input.xlsx
```

### With wasmtime

```bash
# Stdin input (format auto-detected)
wasmtime run exceltocsv.wasm < input.xlsx > output.csv

# Pass flags using -- to separate wasmtime args from module args
wasmtime run exceltocsv.wasm -- -f xls < input.xls > output.csv
wasmtime run exceltocsv.wasm -- --sheet "Sales Q1" < input.xlsx > output.csv

# File input (requires granting directory access)
wasmtime run --dir . exceltocsv.wasm input.xlsx > output.csv
```

## Build

**Prerequisites:** Rust >= 1.88.

**Native binary:**

```bash
cargo build --release
# artifact: target/release/exceltocsv
```

**WASI module** (requires the `wasm32-wasip1` target and `wasm-tools`):

```bash
rustup target add wasm32-wasip1
cargo install wasm-tools

cargo build --target wasm32-wasip1 --release
wasm-tools strip target/wasm32-wasip1/release/exceltocsv.wasm -o exceltocsv.wasm
```

## Testing

```bash
cargo test
```

Tests assert exact byte-for-byte CSV output against files in `tests/expected/`. See [tools/README.md](tools/README.md) for fixture generation instructions.

## Alternatives

- [in2csv](https://csvkit.readthedocs.io/en/latest/scripts/in2csv.html) (Python, part of csvkit) — the CLI that inspired this tool's flag conventions; handles more formats and has richer type inference
- [xlsx2csv](https://github.com/dilshod/xlsx2csv) (Python) — lightweight XLSX-only converter
- [LibreOffice](https://www.libreoffice.org/) / [Gnumeric ssconvert](https://wiki.gnome.org/Projects/Gnumeric/ssconvert) — full office suites with headless conversion; large system dependencies
- [SheetJS](https://sheetjs.com/) — comprehensive JavaScript library; targets browser/Node runtimes, not standalone WASI

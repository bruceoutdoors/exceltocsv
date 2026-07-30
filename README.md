# exceltocsv

CLI to convert Excel files (XLS, XLSX) to CSV. Available as a Linux binary and as a WASI module for embedding in host runtimes.

## Why this exists

Several capable Excel-to-CSV converters already exist. Lightweight options commonly require a language runtime, while office-suite converters are much larger dependencies, and browser-oriented WebAssembly packages are not standalone WASI command modules. exceltocsv packages this narrow conversion as a Linux CLI and WASI module for applications that want to run it within their own resource and capability limits. It is intentionally not a general spreadsheet engine.

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

Download `exceltocsv.wasm` from the [Releases page](../../releases). A `.sha256` checksum file is included.

## Usage

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

Run `exceltocsv --help` for all options.

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

# exceltocsv

CLI to convert Excel files (XLS, XLSX) to CSV. Ships as a Linux binary and as a WASI module for embedding in host runtimes.

- Lightweight, fast, with no runtime dependencies
- Embeddable via WASI — sandbox untrusted workbooks with fine-grained resource and capability limits

## Why this exists

Existing converters (Python's in2csv, xlsx2csv) require a Python runtime. Heavier tools (LibreOffice, Gnumeric's ssconvert) are large system dependencies. Browser-oriented WebAssembly packages target JavaScript runtimes, not WASI. Direct library integration works but produces a one-off binary rather than a reusable conversion boundary. This tool fills the gap: a small, auditable XLS/XLSX-to-CSV converter distributed as a stripped WASI module for sandboxed embedding, with a Linux binary for direct use and scripting.

## Install

### Linux

```sh
curl -fL https://github.com/bruceoutdoors/exceltocsv/releases/latest/download/exceltocsv-linux-amd64 -o exceltocsv
chmod +x exceltocsv
```

Verify the checksum:

```sh
curl -fL https://github.com/bruceoutdoors/exceltocsv/releases/latest/download/exceltocsv-linux-amd64.sha256 | sha256sum -c
```

### macOS and Windows

Build from source — see [Build](#build) below.

### WASI module

Download `exceltocsv.wasm` from the [Releases page](../../releases). A `.sha256` checksum file is included.

## Usage

```bash
# Convert first sheet to CSV on stdout
exceltocsv input.xlsx > output.csv

# Read from stdin (format auto-detected from file magic bytes)
cat input.xlsx | exceltocsv > output.csv
cat input.xls  | exceltocsv -f xls > output.csv

# List all worksheet names
exceltocsv --names input.xlsx

# Select a specific sheet
exceltocsv --sheet "Sales Q1" input.xlsx > sales.csv

# Tab-delimited output
exceltocsv --tabs input.xlsx > output.tsv

# Custom delimiter
exceltocsv -d '|' input.xlsx > output.psv

# Write all sheets to individual CSV files using sheet names
exceltocsv --write-sheets - --use-sheet-names input.xlsx
```

Full option reference:

```text
-f, --format <FMT>     Force format: xls or xlsx (required for stdin if auto-detection fails)
-n, --names            List worksheet names and exit
    --sheet <NAME>     Select worksheet by name (default: first sheet)
    --write-sheets <S> Write sheets to .csv files; - for all, or comma-separated names
    --use-sheet-names  Use sheet names as output filenames (requires --write-sheets)
-d, --delimiter <C>    Output delimiter character (default: comma)
-t, --tabs             Use tab as delimiter
-q, --quotechar <C>    Quote character (default: double-quote)
-u, --quoting <MODE>   minimal | all | nonnumeric | none
-b, --no-doublequote   Disable double-quote escaping (requires --escapechar)
-p, --escapechar <C>   Escape character (used with --no-doublequote)
```

Note: CSV output uses LF line endings (`\n`). RFC 4180 specifies CRLF, but LF is conventional on Unix and avoids test friction.

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

Correctness is verified against known XLS/XLSX fixtures with exact expected CSV outputs.

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

- [in2csv](https://csvkit.readthedocs.io/en/latest/scripts/in2csv.html) (Python, part of csvkit) — the CLI that inspired this tool's flag set; handles more formats and has richer type inference

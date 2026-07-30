# Test Fixture Generation

Fixtures are binary files committed to the repository. Regenerate them only after intentional changes to fixture content.

## XLSX fixtures

Requires Rust and the `generate-fixtures` feature:

```bash
cargo run --bin generate_fixtures --features generate-fixtures
```

This writes `tests/fixtures/simple.xlsx`, `multiple_sheets.xlsx`, `quoted_values.xlsx`, `unicode.xlsx`, and `types.xlsx`.

## XLS fixture

Requires [uv](https://docs.astral.sh/uv/):

```bash
uv run tools/generate_fixtures.py
```

This writes `tests/fixtures/simple.xls`.

## Regenerating expected CSV outputs

After changing fixture content, regenerate the expected outputs:

```bash
cargo build
./target/debug/exceltocsv tests/fixtures/simple.xlsx          > tests/expected/simple.csv
./target/debug/exceltocsv tests/fixtures/simple.xls           > tests/expected/simple_xls.csv
./target/debug/exceltocsv tests/fixtures/multiple_sheets.xlsx > tests/expected/multiple_sheets_sheet1.csv
./target/debug/exceltocsv --sheet Sheet2 tests/fixtures/multiple_sheets.xlsx \
                                                              > tests/expected/multiple_sheets_sheet2.csv
./target/debug/exceltocsv tests/fixtures/quoted_values.xlsx   > tests/expected/quoted_values.csv
./target/debug/exceltocsv tests/fixtures/unicode.xlsx         > tests/expected/unicode.csv
./target/debug/exceltocsv tests/fixtures/types.xlsx           > tests/expected/types.csv
```

Commit both the new fixtures and the updated expected outputs together.

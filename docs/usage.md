# Usage

## Command Line

### Extract Tables

```bash
docx2csv extract [OPTIONS] FILENAME
```

| Option | Short | Default | Description |
|--------|-------|---------|-------------|
| `--format` | `-f` | `csv` | Output format: `csv`, `tsv`, `xlsx`, `xls`, `json` |
| `--singlefile` | `-s` | `false` | Write all tables to a single file |
| `--sizefilter` | `-i` | `0` | Exclude tables with fewer than N rows |
| `--output` | `-o` | (same as input) | Output file path |

### Analyze Document

```bash
docx2csv analyze FILENAME
```

Lists all tables found in the document with their dimensions and styles.

### Examples

```bash
# Extract all tables as separate CSV files
docx2csv extract document.docx

# Extract as XLSX workbook
docx2csv extract document.docx --format xlsx

# Single CSV file with all tables
docx2csv extract document.docx --format csv --singlefile

# Only tables with more than 3 rows
docx2csv extract document.docx --sizefilter 3

# Custom output location
docx2csv extract document.docx --output results/results.csv
```

## Python API

### `extract_tables(filename, strip_space=True)`

Returns a list of table data from a .docx file.

```python
from docx2csv import extract_tables

tables = extract_tables('document.docx')
for table in tables:
    print(f"Table {table['id']}: {table['num_rows']}x{table['num_cols']}")
    for row in table['data']:
        print(row)
```

Each table dict contains:
- `id` - Sequential table number (1-based)
- `num_cols` - Number of columns
- `num_rows` - Number of rows
- `style` - Table style name
- `data` - List of rows, each a list of cell values

### `extract(filename, format='csv', sizefilter=0, singlefile=False, output=None, strip_space=True)`

Extracts tables and saves to file.

```python
from docx2csv import extract

# Save as XLSX
extract('document.docx', format='xlsx', output='output.xlsx')

# Save as single JSON file with metadata
extract('document.docx', format='json', singlefile=True, output='output.json')

# Save as separate CSVs, only tables with 5+ rows
extract('document.docx', format='csv', sizefilter=5)
```

### `analyze(filename)`

Returns table metadata without extracting data.

```python
from docx2csv import analyze

info = analyze('document.docx')
# [{'id': 1, 'num_cols': 3, 'num_rows': 5, 'style': 'Table Grid'}, ...]
```

# docx2csv

Extracts tables from .docx files and saves them as CSV, TSV, XLSX, or JSON.

[![CI](https://github.com/ivbeg/docx2csv/actions/workflows/ci.yml/badge.svg)](https://github.com/ivbeg/docx2csv/actions/workflows/ci.yml)
[![PyPI](https://img.shields.io/pypi/v/docx2csv.svg)](https://pypi.org/project/docx2csv/)
[![Python](https://img.shields.io/pypi/pyversions/docx2csv.svg)](https://pypi.org/project/docx2csv/)

## Installation

```bash
pip install docx2csv
```

For XLS (Excel 97) output support:

```bash
pip install docx2csv[xls]
```

## Quick Start

### Command Line

```bash
# Extract all tables as separate CSV files
docx2csv extract document.docx

# Extract as XLSX
docx2csv extract document.docx --format xlsx

# Extract as single CSV file with all tables
docx2csv extract document.docx --format csv --singlefile

# Filter tables by minimum row count
docx2csv extract document.docx --sizefilter 3

# Specify output location
docx2csv extract document.docx --output results.csv

# Analyze a document (list tables without extracting)
docx2csv analyze document.docx
```

### Python API

```python
from docx2csv import extract_tables, extract, analyze

# Get table data as Python objects
tables = extract_tables('document.docx')
for table in tables:
    print(f"Table {table['id']}: {table['num_rows']} rows x {table['num_cols']} cols")
    for row in table['data']:
        print(row)

# Extract to file
extract('document.docx', format='xlsx', output='output.xlsx')

# Analyze document structure
info = analyze('document.docx')
```

## Output Formats

| Format | Extension | Description |
|--------|-----------|-------------|
| CSV    | `.csv`    | Comma-separated values |
| TSV    | `.tsv`    | Tab-separated values |
| XLSX   | `.xlsx`   | Excel 2007+ (via openpyxl) |
| XLS    | `.xls`    | Excel 97 (requires `pip install docx2csv[xls]`) |
| JSON   | `.json`   | Structured JSON with metadata |

## Requirements

- [click](https://github.com/pallets/click)
- [python-docx](https://github.com/python-openxml/python-docx)
- [openpyxl](https://openpyxl.readthedocs.io/)

## Contributing

Contributions are welcome! Every little bit helps.

1. Fork the repo on GitHub
2. Clone your fork: `git clone https://github.com/your-username/docx2csv.git`
3. Create a branch: `git checkout -b name-of-your-bugfix-or-feature`
4. Make your changes and add tests
5. Run tests: `pytest`
6. Commit and push: `git commit -m "Your message" && git push origin your-branch`
7. Submit a pull request

## Credits

- **Ivan Begtin** - Author and maintainer
- **Vsevolod Oparin** - Optimized table extraction code

## License

BSD 3-Clause License. See [LICENSE](LICENSE) for details.

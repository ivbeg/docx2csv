# Installation

## From PyPI

```bash
pip install docx2csv
```

## With XLS Support

To export to the legacy Excel 97 (.xls) format:

```bash
pip install docx2csv[xls]
```

## From Source

```bash
git clone https://github.com/ivbeg/docx2csv.git
cd docx2csv
pip install -e ".[dev]"
```

## Requirements

- Python >= 3.8
- [click](https://github.com/pallets/click)
- [python-docx](https://github.com/python-openxml/python-docx)
- [openpyxl](https://openpyxl.readthedocs.io/)

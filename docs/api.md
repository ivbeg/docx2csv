# API Reference

## `docx2csv.extract_tables`

```python
extract_tables(filename: str, strip_space: bool = True) -> List[Dict[str, Any]]
```

Extracts all tables from a .docx file.

**Parameters:**
- `filename` (str) - Path to the .docx file
- `strip_space` (bool) - Strip leading/trailing whitespace from cell values (default: `True`)

**Returns:** `List[Dict[str, Any]]` - List of table info dicts with keys:
- `id` (int) - Sequential table number (1-based)
- `num_cols` (int) - Number of columns
- `num_rows` (int) - Number of rows
- `style` (str) - Table style name
- `data` (List[List[str]]) - Table data as list of rows

**Raises:**
- `FileNotFoundError` - If the file does not exist
- `ValueError` - If the path is not a file
- `PackageNotFoundError` - If the file is not a valid .docx

---

## `docx2csv.extract`

```python
extract(
    filename: str,
    format: str = "csv",
    sizefilter: int = 0,
    singlefile: bool = False,
    output: Optional[str] = None,
    strip_space: bool = True,
) -> None
```

Extracts tables from a .docx file and saves them to disk.

**Parameters:**
- `filename` (str) - Path to the .docx file
- `format` (str) - Output format: `"csv"`, `"tsv"`, `"xlsx"`, `"xls"`, or `"json"` (default: `"csv"`)
- `sizefilter` (int) - Exclude tables with fewer than this many rows (default: `0`)
- `singlefile` (bool) - Write all tables to a single file (default: `False`)
- `output` (Optional[str]) - Output file path. Default: same directory as input
- `strip_space` (bool) - Strip leading/trailing whitespace from cell values (default: `True`)

**Raises:**
- `FileNotFoundError` - If the file does not exist
- `ValueError` - If the path is not a file
- `PackageNotFoundError` - If the file is not a valid .docx
- `ImportError` - If XLS format is requested but `xlwt` is not installed

---

## `docx2csv.analyze`

```python
analyze(filename: str) -> List[Dict[str, Any]]
```

Analyzes a .docx file and returns table metadata without extracting data.

**Parameters:**
- `filename` (str) - Path to the .docx file

**Returns:** `List[Dict[str, Any]]` - List of table info dicts with keys:
- `id` (int) - Sequential table number (1-based)
- `num_cols` (int) - Number of columns
- `num_rows` (int) - Number of rows
- `style` (str) - Table style name

**Raises:**
- `FileNotFoundError` - If the file does not exist
- `ValueError` - If the path is not a file
- `PackageNotFoundError` - If the file is not a valid .docx

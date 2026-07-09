# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.3] - 2026-07-09

### Fixed
- TSV output missing UTF-8 encoding that corrupted non-ASCII text
- `--sizefilter` comparing dict length instead of actual row count
- File handles not closed on error (now using context managers)
- `.coveragerc` referencing wrong package name (`budgetlib` -> `docx2csv`)
- README showing wrong command name (`convert` -> `extract`)
- `--singlefile` flag ignored for CSV and TSV formats

### Added
- Input validation with clear error messages for missing/invalid files
- Type hints on all public API functions
- Pytest test suite with 31 tests
- GitHub Actions CI testing Python 3.8-3.12
- `pyproject.toml` (PEP 517/518)
- `CHANGELOG.md` (this file)

### Changed
- Replaced private `python-docx` attribute access with stable public API
- `xlwt` is now an optional dependency (only needed for XLS output)
- Version bump to 0.1.3

### Removed
- Travis CI configuration (replaced by GitHub Actions)
- Sphinx documentation (replaced by Markdown docs)

## [0.1.2] - 2022-08-20

### Added
- `analyze` command
- JSON output format

## [0.1.2] - 2022-08-19

### Changed
- Moved command line to `docx2csv/core.py`
- Added `__main__.py` for `python -m docx2csv`

## [0.1.1] - 2022-01-30

### Fixed
- Requirements in `setup.py` and `requirements.txt`

## [0.1.0] - 2018-01-14

### Added
- First public release on PyPI

[0.1.3]: https://github.com/ivbeg/docx2csv/compare/v0.1.2...v0.1.3
[0.1.2]: https://github.com/ivbeg/docx2csv/compare/v0.1.1...v0.1.2
[0.1.1]: https://github.com/ivbeg/docx2csv/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/ivbeg/docx2csv/releases/tag/v0.1.0

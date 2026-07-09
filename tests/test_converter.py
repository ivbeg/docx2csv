# -*- coding: utf-8 -*-
"""Tests for docx2csv.converter module."""

import csv
import json
import os

import pytest

from docx2csv.converter import extract_tables, extract, analyze


class TestExtractTables:
    """Tests for extract_tables()."""

    def test_returns_correct_count(self, sample_docx):
        tables = extract_tables(sample_docx)
        assert len(tables) == 3

    def test_table_ids_are_sequential(self, sample_docx):
        tables = extract_tables(sample_docx)
        assert tables[0]['id'] == 1
        assert tables[1]['id'] == 2
        assert tables[2]['id'] == 3

    def test_row_counts(self, sample_docx):
        tables = extract_tables(sample_docx)
        assert tables[0]['num_rows'] == 2
        assert tables[1]['num_rows'] == 3
        assert tables[2]['num_rows'] == 1

    def test_col_counts(self, sample_docx):
        tables = extract_tables(sample_docx)
        assert tables[0]['num_cols'] == 2
        assert tables[1]['num_cols'] == 3
        assert tables[2]['num_cols'] == 1

    def test_data_values_simple(self, sample_docx):
        tables = extract_tables(sample_docx)
        assert tables[0]['data'] == [["A", "B"], ["C", "D"]]

    def test_unicode_roundtrip(self, sample_docx):
        """Cyrillic text should be preserved correctly."""
        tables = extract_tables(sample_docx)
        assert tables[1]['data'][0][0] == "Я00"
        assert tables[1]['data'][2][2] == "Я22"

    def test_strip_space_default(self, multline_docx):
        """By default strip_space=True, newlines become spaces."""
        tables = extract_tables(multline_docx)
        assert tables[0]['data'][0][0] == "Line1 Line2 Line3"

    def test_strip_space_false(self, multline_docx):
        """With strip_space=False, newlines become spaces but no stripping."""
        tables = extract_tables(multline_docx, strip_space=False)
        assert tables[0]['data'][0][0] == "Line1 Line2 Line3"

    def test_empty_document(self, empty_docx):
        tables = extract_tables(empty_docx)
        assert tables == []


class TestExtract:
    """Tests for extract()."""

    def test_extract_csv_creates_files(self, sample_docx, tmp_path):
        output_dir = tmp_path / "output"
        output_dir.mkdir()
        extract(sample_docx, format="csv", output=str(output_dir / "out.csv"))
        # singlefile CSV creates one file
        assert (output_dir / "out.csv").exists()

    def test_extract_csv_separate_files(self, sample_docx, tmp_path):
        output_dir = tmp_path / "output"
        output_dir.mkdir()
        extract(sample_docx, format="csv", singlefile=False)
        base = sample_docx.rsplit(".", 1)[0]
        # 3 tables → 3 CSV files
        assert os.path.exists(base + "_1.csv")
        assert os.path.exists(base + "_2.csv")
        assert os.path.exists(base + "_3.csv")

    def test_extract_tsv_utf8_encoding(self, sample_docx, tmp_path):
        """SPEC-001: TSV output must be valid UTF-8."""
        output = str(tmp_path / "out.tsv")
        extract(sample_docx, format="tsv", singlefile=True, output=output)
        with open(output, "rb") as f:
            raw = f.read()
        # Verify it's valid UTF-8
        raw.decode("utf-8")
        # Verify Cyrillic content is present
        assert "Я00" in raw.decode("utf-8")

    def test_extract_csv_utf8_encoding(self, sample_docx, tmp_path):
        """CSV output must be valid UTF-8."""
        output = str(tmp_path / "out.csv")
        extract(sample_docx, format="csv", singlefile=True, output=output)
        with open(output, "rb") as f:
            raw = f.read()
        raw.decode("utf-8")
        assert "Я00" in raw.decode("utf-8")

    def test_extract_json_valid(self, sample_docx, tmp_path):
        output = str(tmp_path / "out.json")
        extract(sample_docx, format="json", singlefile=True, output=output)
        with open(output, "r", encoding="utf-8") as f:
            data = json.load(f)
        assert data['num_tables'] == 3
        assert data['filename'] == sample_docx
        assert 'timestamp' in data
        assert len(data['tables']) == 3

    def test_extract_xlsx_creates_file(self, sample_docx, tmp_path):
        output = str(tmp_path / "out.xlsx")
        extract(sample_docx, format="xlsx", output=output)
        assert os.path.exists(output)

    def test_sizefilter_excludes_small_tables(self, sample_docx, tmp_path):
        """SPEC-002: sizefilter=3 should exclude tables with < 3 rows."""
        base = sample_docx.rsplit(".", 1)[0]
        extract(sample_docx, format="csv", sizefilter=3, singlefile=False)
        # Only table 2 (3 rows) passes; tables 1 (2 rows) and 3 (1 row) excluded
        # The single passing table is renumbered as _1.csv
        assert os.path.exists(base + "_1.csv")
        with open(base + "_1.csv", "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            rows = list(reader)
        assert len(rows) == 3
        assert rows[0][0] == "Я00"

    def test_singlefile_csv_one_file(self, sample_docx, tmp_path):
        """SPEC-007: singlefile CSV produces one file."""
        output = str(tmp_path / "combined.csv")
        extract(sample_docx, format="csv", singlefile=True, output=output)
        with open(output, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            rows = list(reader)
        # 3 tables: 2 rows + blank + 3 rows + blank + 1 row = 8 rows
        assert len(rows) == 8  # 2 + 1(blank) + 3 + 1(blank) + 1

    def test_singlefile_tsv_one_file(self, sample_docx, tmp_path):
        """SPEC-007: singlefile TSV produces one file."""
        output = str(tmp_path / "combined.tsv")
        extract(sample_docx, format="tsv", singlefile=True, output=output)
        assert os.path.exists(output)


class TestAnalyze:
    """Tests for analyze()."""

    def test_returns_metadata(self, sample_docx):
        info = analyze(sample_docx)
        assert len(info) == 3
        assert info[0]['id'] == 1
        assert info[0]['num_cols'] == 2
        assert info[0]['num_rows'] == 2
        assert 'style' in info[0]

    def test_empty_document(self, empty_docx):
        info = analyze(empty_docx)
        assert info == []


class TestValidation:
    """Tests for input validation."""

    def test_missing_file_raises_file_not_found(self):
        with pytest.raises(FileNotFoundError):
            extract_tables("/nonexistent/path/file.docx")

    def test_directory_raises_value_error(self, tmp_path):
        with pytest.raises(ValueError, match="Path is not a file"):
            extract_tables(str(tmp_path))

    def test_non_docx_extension_warns(self, tmp_path):
        fake_file = tmp_path / "test.txt"
        fake_file.write_text("not a docx")
        # Warning fires, then Document() raises PackageNotFoundError
        with pytest.warns(UserWarning, match="does not have .docx"):
            with pytest.raises(Exception):
                extract_tables(str(fake_file))

    def test_extract_missing_file_error(self, tmp_path):
        """SPEC-006: extract() raises FileNotFoundError for missing file."""
        with pytest.raises(FileNotFoundError):
            extract("/nonexistent/file.docx")

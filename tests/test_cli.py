# -*- coding: utf-8 -*-
"""Tests for the Click CLI (docx2csv.core)."""

import pytest
from click.testing import CliRunner

from docx2csv.core import cli


class TestCLIExtract:
    """Tests for the CLI extract command."""

    def test_extract_help(self):
        runner = CliRunner()
        result = runner.invoke(cli, ["extract", "--help"])
        assert result.exit_code == 0
        assert "--format" in result.output
        assert "--singlefile" in result.output
        assert "--sizefilter" in result.output
        assert "--output" in result.output

    def test_extract_missing_file_error(self):
        """SPEC-006: CLI shows clean error for missing file."""
        runner = CliRunner()
        result = runner.invoke(cli, ["extract", "/nonexistent/file.docx"])
        assert result.exit_code == 1
        assert "Error" in result.output
        assert "File not found" in result.output

    def test_extract_directory_error(self, tmp_path):
        """SPEC-006: CLI shows clean error for directory path."""
        runner = CliRunner()
        result = runner.invoke(cli, ["extract", str(tmp_path)])
        assert result.exit_code == 1
        assert "Error" in result.output

    def test_extract_success(self, sample_docx):
        """CLI extract creates output files."""
        runner = CliRunner()
        result = runner.invoke(cli, ["extract", sample_docx, "--format", "csv"])
        assert result.exit_code == 0


class TestCLIAnalyze:
    """Tests for the CLI analyze command."""

    def test_analyze_help(self):
        runner = CliRunner()
        result = runner.invoke(cli, ["analyze", "--help"])
        assert result.exit_code == 0

    def test_analyze_missing_file_error(self):
        """SPEC-006: CLI analyze shows clean error for missing file."""
        runner = CliRunner()
        result = runner.invoke(cli, ["analyze", "/nonexistent/file.docx"])
        assert result.exit_code == 1
        assert "Error" in result.output

    def test_analyze_success(self, sample_docx):
        """CLI analyze runs without error."""
        runner = CliRunner()
        result = runner.invoke(cli, ["analyze", sample_docx])
        assert result.exit_code == 0

#!/usr/bin/env python
# -*- coding: utf8 -*-

import click

import docx2csv


@click.group()
def cli1():
    """Extracts tables from DOCX files as CSV or XLSX.
        Use command: "docx2csv convert <filename>" to run extraction.
        It will create files like filename_1.csv, filename_2.csv for each table found.

    """
    pass


@cli1.command()
@click.option('--format', default='csv', help='Output format: CSV, TSV, XLSX')
@click.option('--singlefile', default=False, help='Outputs XLS file with multiple sheets' )
@click.option('--sizefilter', default=0, help='Filters table by size number of rows')
@click.option('--output', default=None, help='Choose location of output file, default same location as input')
@click.argument('filename')
def extract(format, sizefilter, singlefile, filename, output):
    """
        docx to csv convertor (http://github.com/ivbeg/docx2csv)

        Extracts tables from DOCX files as CSV or TSV or XLSX.

        Use command: "docx2csv convert <filename>" to run extraction.
        It will create files like filename_1.csv, filename_2.csv for each table found.
    """
    docx2csv.extract(filename, format, sizefilter, singlefile, output)


cli = click.CommandCollection(sources=[cli1])

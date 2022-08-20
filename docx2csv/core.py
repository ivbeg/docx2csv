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
@click.argument('filename')
@click.option('--format', '-f', default='csv', help='Output format: CSV, TSV, XLSX')
@click.option('--singlefile', '-s', is_flag=True, show_default=True, default=False, help='Outputs XLS file with multiple sheets' )
@click.option('--sizefilter', '-i', default=0, help='Filters table by size number of rows')
@click.option('--output', '-o', default=None, help='Choose location of output file, default same location as input')
def extract(filename, format, sizefilter, singlefile, output):
    """
        Extracts tables from DOCX files as CSV or TSV or XLSX.

        Use command: "docx2csv extract <filename>" to run extraction.
        It will create files like filename_1.csv, filename_2.csv for each table found.
    """
    docx2csv.extract(filename, format, sizefilter, singlefile, output)


@click.group()
def cli2():
    """Analyses of DOCX file, lists all existing tables
    """
    pass


@cli2.command()
@click.argument('filename')
def analyze(filename):
    """
        Analyzes .docx file and finds tables
    """
    from pprint import pprint
    tableinfo = docx2csv.analyze(filename)
    pprint(tableinfo)



cli = click.CommandCollection(sources=[cli1, cli2])

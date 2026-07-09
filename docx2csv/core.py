#!/usr/bin/env python
# -*- coding: utf8 -*-

import sys

import click

import docx2csv
from docx.opc.exceptions import PackageNotFoundError


@click.group()
def cli1():
    """Extracts tables from DOCX files as CSV or XLSX.

        Use command: "docx2csv extract <filename>" to run extraction.
        It will create files like filename_1.csv, filename_2.csv for each table found.

    """
    pass


@cli1.command()
@click.argument('filename')
@click.option('--format', '-f', default='csv', help='Output format: CSV, TSV, XLSX')
@click.option('--singlefile', '-s', is_flag=True, show_default=True, default=False, help='Outputs single file with multiple tables')
@click.option('--sizefilter', '-i', default=0, help='Filters table by minimum number of rows')
@click.option('--output', '-o', default=None, help='Choose location of output file, default same location as input')
def extract(filename, format, sizefilter, singlefile, output):
    """
        Extracts tables from DOCX files as CSV or TSV or XLSX.

        Use command: "docx2csv extract <filename>" to run extraction.
        It will create files like filename_1.csv, filename_2.csv for each table found.
    """
    try:
        docx2csv.extract(filename, format, sizefilter, singlefile, output)
    except FileNotFoundError as e:
        click.echo("Error: %s" % e, err=True)
        sys.exit(1)
    except ValueError as e:
        click.echo("Error: %s" % e, err=True)
        sys.exit(1)
    except PackageNotFoundError:
        click.echo("Error: '%s' is not a valid .docx file." % filename, err=True)
        sys.exit(1)


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
    try:
        tableinfo = docx2csv.analyze(filename)
        pprint(tableinfo)
    except FileNotFoundError as e:
        click.echo("Error: %s" % e, err=True)
        sys.exit(1)
    except ValueError as e:
        click.echo("Error: %s" % e, err=True)
        sys.exit(1)
    except PackageNotFoundError:
        click.echo("Error: '%s' is not a valid .docx file." % filename, err=True)
        sys.exit(1)


cli = click.CommandCollection(sources=[cli1, cli2])

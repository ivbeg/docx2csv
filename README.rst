============
Command line
============

Usage: docx2csv [OPTIONS] FILENAME

  docx to csv convertor (http://github.com/ivbeg/docx2csv)
  Extracts tables from DOCX files as CSV or XLSX.

  Use command: "docx2csv convert <filename>" to run extraction. It will
  create files like filename_1.csv, filename_2.csv for each table found.

Options:
  --format TEXT         Output format: CSV, XLSX
  --singlefile TEXT     Outputs single XLS file with multiple sheets: True or False
  --sizefilter INTEGER  Filters table by size number of rows
  --help                Show this message and exit.

Examples
========
docx2csv --format csv --sizefilter 3 CP_CONTRACT_160166.docx

Extracts tables from file CP_CONTRACT_160166.docx with number of rows > 3 and
saves results as CSV files.


Code
====


Function "extract_tables" returns list of tables from docx file and function "extract" extracts tables as xlsx, xls or csv file. If 'csv' file selected you need to specify single file csv or multiple files
    >>> from docx2csv import extract_tables, extract
    >>> tables = extract_tables('some_file.docx')

    returns list of tables
    >>> extract(filename='some_file.docx', format="xlsx", output='some_file.xlsx')
    saves all tables from some_file.docx to some_file.xlsx

    >>> extract(filename='some_file.docx', format="csv", singlefile=False)
    saves all tables from some_file.docx to some_file_1.csv, some_file_2.csv and etc.



Requirements
============
* click https://github.com/pallets/click
* xlwt https://github.com/python-excel/xlwt
* python-docx https://github.com/python-openxml/python-docx
* openpyxl https://bitbucket.org/openpyxl/openpyxl/src


Acknowledgements
================
Thanks to Vsevolod Oparin (https://www.facebook.com/vsevolod.oparin) for optimized "extract_table" code

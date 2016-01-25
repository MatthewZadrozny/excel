# Spreadsheets

## Utilities for converting .ods files to sqlite3 databases and html tables

Can be imported or used on the command line, with optional arguments to specify sheets. 

To convert an ods spreadsheet comprised of sheets A, B, C, and D to a SQLite3 database with only sheets A, B, C run:

`$ python convert.py EXAMPLE.ods A B C`


To convert the first sheet in an ods spreadsheet to an html table (sortable with sortable.js), run:

`$ python convert.py EXAMPLE.ods`


### Dependencies

* [Sortable] (http://www.kryogenix.org/code/browser/sorttable/)
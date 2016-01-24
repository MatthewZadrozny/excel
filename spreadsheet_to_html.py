#!/bin/python

# create html table(s) from .ods file
# Author(s): Matthew Zadrozny <m.zadrozny@gmail.com>
# Downloaded 2015-10-30 from 
# http://www.ctroms.com/blog/code/python/2011/04/20/csv-to-html-table-with-python/ 
# modified to work with .ods files
# Version 2 - added css class to all columns except header
# Uses http://www.kryogenix.org/code/browser/sorttable/

import codecs
import ezodf
import ods
import sys


def convert(ods_file, sheets=None):
    spreadsheet = ezodf.opendoc(ods_file)
    sheets = [s.name for s in spreadsheet.sheets]
    table = spreadsheet.sheets[sheets[0]] # temporary, remove
    rows = list(table.rows())
    reader = [[c.value for c in r] for r in rows]

    htmlfile = open(str(ods_file).replace('ods', 'html'), 'w')
    htmlfile.write('<meta charset="utf-8" />')
    htmlfile.write('<script src="sorttable.js"></script>')
    htmlfile.write('<table class="sortable" border="1" style="width:100%">\n')

    for row in reader:
        htmlfile.write('</tr>\n')
        for column in row:
            print column
            if column == None: 
                column = str(None)
            elif type(column) == float:
                column = str(None)
            else:
                column = column.encode('utf-8')
            htmlfile.write('<th>' + column + '</th>\n')
        htmlfile.write('</tr>\n')

    htmlfile.write('</table>\n')

    print "Done!"

    htmlfile.close()


if __name__ == '__main__':
    spreadsheet_file = sys.argv[1]
    sheets = sys.argv[2:] # Lack of arguments means len(sheets) == 0
    convert(spreadsheet_file, sheets)
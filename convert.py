#!/usr/bin/python3
# -*- coding: utf-8 -*-

# Converts .ods to: 
#	array, 
# 	.txt, 
# 	.html, 
#	pandas DataFrame
# 	sqlite3


import ezodf
import numpy as np
import pandas as pd
import sqlite3
import sys


def rename(ods, extension):
    '''Replaces .ods extension with new extension / path'''
    return ods[:-3] + extension



def array(ods, sheet):
    '''Converts individual .ods spreadsheet sheet to an array'''
    spreadsheet = ezodf.opendoc(ods)
    table = spreadsheet.sheets[sheet]
    return [[str(c.value) 
             if str(c.value)[-2:] != '.0' 
             else str(c.value)[:-2] for c in r] 
             for r in list(table.rows())]



def txt(ods, sheet):
    '''Converts .ods to .txt'''
    arr = array(ods, sheet)

    # Add pretty formatting
    # contents = str(["%-40s %-25s %-10s" % (x[0], x[1], x[2])][0])

    output = '\n'.join(['\t'.join(row) for row in arr])

    newfile = rename(ods, "txt")

    with open(newfile, "w") as f:
        f.write(output +'\n')
    


def html(ods, sheet):
	'''Converts .ods to .html (table)'''
    arr = array(ods, sheet)

    with open(rename(ods, 'html'), 'w') as htmlfile:

        htmlfile.write('<meta charset="utf-8" />')
        htmlfile.write('<script src="sorttable.js"></script>')   
        htmlfile.write('<table class="sortable" border="1" style="width:100%">\n')

        for row in arr:
            htmlfile.write('</tr>\n') # Should this be an open tag?
            for cell in row:
                htmlfile.write('<th>' + cell + '</th>\n')
            htmlfile.write('</tr>\n')

        htmlfile.write('</table>\n')



def pandas(ods, sheet, dropped=True):
	'''Converts an .ods sheet to a pandas DataFrame'''
	arr = array(ods, sheet)
	df = pd.DataFrame(arr[1:], columns=arr[0])
	if dropped:
		return df.replace("None", float("nan"))\
				 .dropna(how='all')\
				 .dropna(axis=1, how='all')
	else:
		return df



def sqlite(ods, sheets=None):
    '''Converts .ods into an sqlite3 database'''

    db = rename(ods, "db")
    conn = sqlite3.connect(db) 
    c = conn.cursor()

    spreadsheet = ezodf.opendoc(ods)

    # Convert every sheet
    if sheets == None or len(sheets) == 0: # len == 0 for command line
        sheets = [s.name for s in spreadsheet.sheets]
        

    for sheet in sheets:

        table = spreadsheet.sheets[sheet]
 
        sheet_name = "\'"+table.name+"\'" # Prevents keyword and spacing errors

        rows = list(table.rows())
        
        header = [str(column.value) 
                    if str(column.value) != 'None' else "COLUMN_"+str(i+1) 
                    for i, column in enumerate(rows[0])]

        create_command = '''create table %s %s''' % (sheet_name, tuple(header))

        c.execute(create_command)

        for row in rows[1:]:

            # Floats as PK's going to be a problem?
            # Temporarily converting for quotes app

            # https://docs.python.org/2/library/sqlite3.html#introduction
            # http://pythonhosted.org/ezodf/tableobjects.html#cell-class
            
            t = tuple([int(cell.value) if cell.value_type=='float' else unicode(cell.value) for cell in row])

            print t[0]

            marks = ', '.join(list(len(header)*"?"))

            command = '''insert into %s values (%s)''' % (sheet_name, marks)

            c.execute(command, t)

    
    conn.commit()
    c.close()




if __name__ == '__main__':
    output = sys.argv[1] 
    ods = sys.argv[2]
    # Change back to 3: so as to handle multiple sheets
    sheets = sys.argv[3] # Lack of arguments means len(sheets) == 0

    if output == 'txt':
        txt(ods, sheets)
    elif output == 'html':   	
        html(ods, sheets)
    elif output == 'sql':
        sql(ods, sheets)

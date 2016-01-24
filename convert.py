#!/usr/bin/python3
# -*- coding: utf-8 -*-

import ezodf
import sqlite3
import sys


# Write a line to backup the file before working on it each time

# Fix the PK keys so that they're not floats!

def convert(ods_file, sheets=None):
    '''Takes an ODS spreadsheet and optional 
    list of sheets and turns it into a DB'''

    db_name = str(ods_file)[0:-4]+'.db' 
    conn = sqlite3.connect(db_name) 
    c = conn.cursor()

    spreadsheet = ezodf.opendoc(ods_file)

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

        # Necessary ?
        # Make this work for irregularly formatted spreadhseets
        # maximum = sorted([len(row) for row in rows])[-1]
        
        # if maximum > len(header):
        #   addendum = ["Column_"+str(x) for x in range(len(header)+1, maximum+2)]

        #   header = header.extend()

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
    spreadsheet_file = sys.argv[1]

    sheets = sys.argv[2:] # Lack of arguments means len(sheets) == 0

    convert(spreadsheet_file, sheets)
import sqlite3, xlrd 
import create #My code

f = "C:\Users\MZ\Dropbox\AGENDA\QUOTATIONS.xlsx"

with xlrd.open_workbook(f, encoding_override="cp1252") as workbook: 
	connection, cursor = create.create_database(f)		
	
	for sheet_name in workbook.sheet_names():
		columns, header = create.create_table(cursor, sheet_name, workbook)
		create.populate_table(columns, cursor, sheet_name, workbook)
	
	connection.commit()
	connection.close()
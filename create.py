import codecs, sqlite3, xlrd 

def create_database(f):
	'''Takes an excel spreadsheet f, creates a database of the same name, 
	and returns a connection and a cursor object.'''
	database = f.split(".")[0]+".db" #Removes extension, adds db extension
	connection = sqlite3.connect(database)
	connection.text_factory = str
	cursor = connection.cursor()
	return connection, cursor


def create_table(cursor, sheet, workbook):
	'''Creates a table from a worksheet and returns a header.'''
	cursor.execute('DROP TABLE IF EXISTS ' + sheet) #Where each sheet = a table	
	sheet_object = workbook.sheet_by_name(sheet)
	header = sheet_object.row_values(0)[1:] #Skip first column in Quotations (my keys)
	columns = header[:] #A list of the objects in the header
	for i, column in enumerate(header): header[i] = column + ' text' 
	header = ', '.join(header) 
	cursor.execute('CREATE TABLE '+ sheet +  '(' + header + ')')
	return columns, header


def populate_table(columns, cursor, sheet, workbook):
	'''Populates a table with the contents of its corresponding Excel sheet.'''
	sheet_object = workbook.sheet_by_name(sheet)
	vals = ("?, " * (len(columns) - 1)) + "?" #Subtract 1 so as not to have a comma at end
	header = ", ".join(columns)
	statement = 'INSERT INTO '+sheet+'('+header+')'+'VALUES'+'('+vals+')'
	for rownum in range(1, sheet_object.nrows):
		row = sheet_object.row_values(rownum)
		row = row[1:] #Skip over my primary keys
		row = (tuple(row))
		cursor.execute(statement, row) 
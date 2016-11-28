import mysql.connector
import xlsxwriter

input_conf = {
	'columns': '*', #csv column names
	'table': 'chainListImportant',
	'condition': "name=\"7-Eleven\"",
	"extra": "",
	"connectionConfig": {
		'user': 'root',
	        'password': 'home123',
        	'host': '127.0.0.1',
	        'database': 'business',
	        'raise_on_warnings': True,
	}
}

output_conf = {
	'type': 't', # t for table or e for excel
	'fileName': 'output.xlsx', # file name if using excel
	'table': 'chainListNew',
	"connectionConfig": {
                'user': 'root',
                'password': 'home123',
                'host': '127.0.0.1',
                'database': 'business',
                'raise_on_warnings': True,
        }
}


def read_data():
	cnx = mysql.connector.connect(**input_conf["connectionConfig"])
	
	cursor = cnx.cursor(buffered=True)

	columns = input_conf['columns']

	# TODO find a better way to know all column names
	if columns == '*':
		query = "SELECT * FROM " + input_conf["table"] + " LIMIT 1"
		cursor.execute(query)
		columns = ",".join(list(cursor.column_names))
	
	query = "SELECT " + columns + " FROM " + input_conf["table"]


	if len(input_conf["condition"]):
		query += " WHERE " + input_conf["condition"]

	query += " " + input_conf["extra"]

	cursor.execute(query)
	
	rows = [item for item in cursor]

	cursor.close()
	
	cnx.close()
	
	return rows, columns


def write_data(rows, columns):
	if output_conf["type"] == 't':
		write_table(rows, columns)
	else:
		write_excel(rows, columns)


def write_table(rows, columns):
	cnx = mysql.connector.connect(**output_conf["connectionConfig"])
       
        cursor = cnx.cursor(buffered=True)

	table = output_conf["table"]
 
        query = ("INSERT INTO " + table + " (" + columns + ") VALUES (" + (",".join(["%s" for __ in xrange(len(columns.split(",")))])) + ")")

	for row in rows:
		cursor.execute(query, row)

	cnx.commit()

	cursor.close()
	cnx.close()


def write_excel(rows, columns):
	workbook = xlsxwriter.Workbook(output_conf['fileName'])
	worksheet = workbook.add_worksheet()
	
	columns = columns.split(",")

	for row in xrange(1, len(rows) + 1):
		for i in xrange(len(columns)):
			worksheet.write(0, i, columns[i])
			worksheet.write(row, i, rows[row - 1][i])
	

	workbook.close()
	

if __name__ == "__main__":
	rows, columns = read_data()

	write_data(rows, columns)

	print "\nTransfered " + str(len(rows)) + " rows.\n"

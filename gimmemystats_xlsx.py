import psycopg2
import psycopg2.extras
import xlsxwriter
from datetime import datetime
from dbconf import *


def gimme_nao():
	conn = psycopg2.connect("""
		dbname='{}'
		user='{}'
		host='{}'
		password='{}'
		""".format(dbname, user, host, passwd))

	# Defined cursor to work with
	cur = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)

	report_date_str = '{}-{}'.format(*report_date[:-1])
	print('Generating reports for {} started...\n'.format(report_date_str))

	# Executing every queryset from the querysets list
	for table_name, qs in querysets.items():

		cur.execute(qs)

		# col_style = xlwt.easyxf('font: bold on')
		filename = '{}_{}.xlsx'.format(table_name, report_date_str)

		wb = xlsxwriter.Workbook(filename)
		ws = wb.add_worksheet(table_name)

		# workbook style
		bold = wb.add_format({'bold': True})

		# Creating column headers
		colnames = [desc[0] for desc in cur.description]

		for colnum, colname in enumerate(colnames):
			ws.set_column(colnum, colnum, len(str(colname)))
			ws.write(0, colnum, colname, bold)

		# Extracting data from the DB, and then 
		# writing it to spreadsheet
		rows = cur.fetchall()

		count_rows = 0
		date_format = wb.add_format({'num_format': 'dd/mm/yyyy'})
		for rownum, row in enumerate(rows):

			longest_cell = 10
			for colnum, cell in enumerate(row):

				if type(cell) is datetime:
					# ws.write(rownum+1, colnum, str(cell)[:19])
					ws.write_datetime(rownum+1, colnum, cell, date_format)
				else:
					ws.write(rownum+1, colnum, cell)


				# TODO: Adjusting column width to widest data found
				if ws.col_sizes.get(colnum) < len(str(cell)):
					ws.set_column(colnum, colnum, len(str(cell)))


			count_rows += 1
		
		print('-' * 60)
		print('-> report: {}'.format(filename))
		print('\tTotal: {} rows'.format(count_rows))
		print('-' * 60, '\n')

		wb.close()


try:
	gimme_nao()
except psycopg2.OperationalError as e:
	print(e)

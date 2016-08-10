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
		filename = 'report-{}_{}.xlsx'.format(table_name, report_date_str)

		wb = xlsxwriter.Workbook(filename)
		ws = wb.add_worksheet(table_name)

		# Creating column headers
		colnames = [desc[0] for desc in cur.description]

		for colnum, colname in enumerate(colnames):
			ws.write(0, colnum, colname) #, col_style)

		# Extracting data from the DB, and then 
		# writing it to spreadsheet
		rows = cur.fetchall()

		count_rows = 0
		for rownum, row in enumerate(rows):

			for colnum, cell in enumerate(row):

				if type(cell) is datetime:
					ws.write(rownum+1, colnum, str(cell)[:19])
				else:
					ws.write(rownum+1, colnum, cell)

				# TODO: Adjusting column width to widest data found

			count_rows += 1

			# if count_rows == 65535:
			# 	print('[Error] Too many freaking rows in {}! Repors truncated!\n'.format(table_name))
			# 	break
		
		print('-' * 60)
		print('-> report: {}'.format(filename))
		print('\tTotal: {} rows'.format(count_rows))
		print('-' * 60, '\n')

		wb.close()


try:
	gimme_nao()
except psycopg2.OperationalError as e:
	print(e)

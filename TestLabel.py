
import win32api
import win32print

import pyodbc

from fpdf import FPDF   

conn = pyodbc.connect('DSN=QDSN_10.0.0.1;UID=username;PWD=password')
cursor = conn.cursor()
cursor.execute("select ODITEM, ODQORD from HDSDATA.OEORDT where ODORD# = 12345") # pull sample order

parts_list = []													# [part number, quantity]

for rows in cursor:
	parts_list.append([rows.ODITEM.rstrip(), int(rows.ODQORD)])

label = FPDF(orientation='P', unit='pt', format=(144,25.2))		# 2" x 0.35"
label.set_margins(0, 0)											# we don't want margins
label.set_auto_page_break(0)									# turn off page breaks
label.set_font('Courier', 'B', 18)

for p in parts_list:
	for i in range(0, p[1]):
		label.add_page()
		label.cell(144, 25, p[0], 0, 0, 'C')

label.output('TEMP_LABEL.pdf')

# automatically print the label
#win32api.ShellExecute (0, "printto", 'TEMP_LABEL.PDF', '\\\\printserver\\hp6100', ".", 0)

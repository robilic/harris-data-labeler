from Tkinter import *
import tkFont

import win32api
import win32print

import pyodbc

from fpdf import FPDF   

class App:

	def __init__(self, master):

		frame = Frame(master)
		frame.pack()
		
		arial18 = tkFont.Font(family='Arial', size=18, weight='bold')  # big font so we can read it in the shop

		self.order_label = Label(frame, text="ORDER #:", font=arial18)
		self.order_label.grid(row=0, column=0)

		self.order_value = Entry(frame, text="", font=arial18)
		self.order_value.grid(row=0, column=1, columnspan=2)

		self.generate_button = Button(frame, text="Generate Labels", font=arial18, command=self.generate_labels)
		self.generate_button.grid(row=0, column=3)
		
		self.single_button = Button(frame, text="Print One", font=arial18, command=self.print_one_label)
		self.single_button.grid(row=2, column=3)

		self.single_label = Label(frame, text="Single Label:", font=arial18)
		self.single_label.grid(row=2, column=0)

		self.single_value = Entry(frame, text="", font=arial18)
		self.single_value.grid(row=2, column=1, columnspan=2)

		self.quit_button = Button(frame, text="QUIT", font=arial18, fg="red", command=frame.quit)
		self.quit_button.grid(row=11, column=3)
	
	def generate_labels(self):
		conn = pyodbc.connect('DSN=QDSN_10.0.0.1;UID=username;PWD=password')
		cursor = conn.cursor()
		cursor.execute("select ODITEM, ODQORD from HDSDATA.OEORDT where ODORD# = " + self.order_value.get() )

		parts_list = []													# [part number, quantity]

		for rows in cursor:
			parts_list.append([rows.ODITEM.rstrip(), int(rows.ODQORD)])

		# 72 pts in an inch
		l_width  = 162
		l_height = 54
		label = FPDF(orientation='P', unit='pt', format=(l_width, l_height) )
		label.set_margins(0, 0)											# we don't want margins
		label.set_auto_page_break(0)									# turn off page breaks
		label.set_font('Courier', 'B', 22)

		for p in parts_list:
			for i in range(0, p[1]):
				label.add_page()
				label.cell(l_width, l_height, p[0], 0, 0, 'C')

		label.output(temp_pdf_file)
		# automatically print the label
		#win32api.ShellExecute (0, "printto", temp_pdf_file, label_printer, ".", 0)

	def print_one_label(self):
		# no database query, just print a label for a manually entered part #
		# 72 pts in an inch
		l_width  = 162
		l_height = 54
		label = FPDF(orientation='P', unit='pt', format=(l_width, l_height) )
		label.set_margins(0, 0)											# we don't want margins
		label.set_auto_page_break(0)									# turn off page breaks
		label.set_font('Courier', 'B', 22)
		label.add_page()
		label.cell(l_width, l_height, self.single_value.get(), 0, 0, 'C')

		label.output(temp_pdf_file)

root = Tk()
root.wm_title("Parts Labeler")

app = App(root)

label_printer = 'ZDesigner GC420d (EPL)' # \\\\printserver\\zebra'
temp_pdf_file = 'TEMP_LABEL.PDF'

root.mainloop()
root.destroy()

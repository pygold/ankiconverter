
import json
import AnkiTools
from AnkiTools import anki_convert
import os
import sys, getopt
from openpyxl import Workbook, load_workbook
import shutil
import xml.sax.saxutils as saxutils



class AnkiConverter():
	def __init__(self, inputfile, filename):
		self.inputfile = inputfile
		self.outputfile = filename + ".xlsx"
		self.filename = filename

	def import_file(self):
		print(self.inputfile)

	def run(self):
		print(self.outputfile)
		print(self.inputfile)

		try:
			anki_convert(self.inputfile, self.outputfile)
			self.get_parser_from(self.outputfile)
		except Exception as e:
			print(repr(e))

	def get_parser_from(self, outputfile):
		wb = load_workbook(outputfile)

		print(wb.sheetnames)
		sheet = wb.get_sheet_by_name(wb.sheetnames[2])
		for row in sheet.rows:
			for cell in row:
				value = cell.value
				if value:
					decoded = saxutils.escape(str(value))
					with open("{}.txt".format(self.filename), 'a+', encoding="utf-8") as f:
						f.write("\n" + decoded)

	# input_filename = "data/test.apkg"
	# output_filename = "data/test.xlsx"
	# ankiconvert = AnkiConverter(input_filename, output_filename)
	# anki_convert(input_filename, out_file=output_filename)
	# anki_convert('my_workbook.xlsx', out_format='.apkg')

def main(argv):
	inputfile   = ''
	
	try:
		inputfile = argv[1]
		extension = inputfile.split(".")[1]
		filename = inputfile.split(".")[0]
	except:
		print ('AnkiToExcel.py <inputfile.apkg>')
		exit(2)
	
	if extension in ['apkg']:
		ankitoexcel = AnkiConverter(inputfile, filename)
		ankitoexcel.run()
	else:
		print ('AnkiToExcel.py <inputfile.apkg>')
		exit(2)

if __name__ == "__main__":
	main(sys.argv)
	
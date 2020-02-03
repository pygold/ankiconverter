
import json
import AnkiTools
from AnkiTools import anki_convert
import os
import sys, getopt
from openpyxl import Workbook, load_workbook
import shutil
import xml.sax.saxutils as saxutils



class AnkiConverter():
	def __init__(self, inputfile, outputfile):
		self.inputfile = inputfile
		self.outputfile = outputfile

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
		sheet = wb.get_sheet_by_name("Cloze-c6a07")
		for row in sheet.rows:
			for cell in row:
				value = cell.value
				if value:
					decoded = saxutils.escape(str(value))
					with open("sheet_date.txt", 'a+') as f:
						f.write("\n" + decoded)

	def saveAs(self):
		self.file_headers = self.get_parser_from(outputfile)
		self.file_contentes = self.get_parser_froM("content")

		if self.file_contentes:
			new_workbook = Workbook
			new_sheet = new_workbook.active()
			try:
				KeyboardInterrupt
			except Exception as e:
				print(repr(e))

			try:
				self.workbook_content = self.find_all_tags(self.rows)

			except:
				print("Can not find any tags in workbook Please retry again")
			
			for content in self.workbook_content:
				if content:
					for tag in contents:
						drp
				else:
					continue

	# input_filename = "data/test.apkg"
	# output_filename = "data/test.xlsx"
	# ankiconvert = AnkiConverter(input_filename, output_filename)
	# anki_convert(input_filename, out_file=output_filename)
	# anki_convert('my_workbook.xlsx', out_format='.apkg')

def main(argv):
	# inputfile   = ''
	# outputfile  = ''
	# try:
	# 	opts, args = getopt.getopt(argv, "hi:o:", ["ifile=","ofile="])
	# except getopt.GetoptError:
	# 	print ('ankitoexcel.py -i <inputfile> -o <outputfile>')
	# 	sys.exit(2)

	# for opt, arg in opts:
	# 	print(opt, arg)
	# 	if opt == '-h':
	# 		print ('test.py -i <inputfile> -o <outputfile>')
	# 		sys.exit()
	# 	elif opt in ("-i", "--ifile"):
	# 		inputfile = arg
	# 	elif opt in ("-o", "--ofile"):
	# 		outputfile = arg

	inputfile = "test.apkg"
	outputfile = "text.xlsx"

	if inputfile and outputfile:
		ankitoexcel = AnkiConverter(inputfile, outputfile)
		ankitoexcel.run()

	else:
		print ('ankitoexcel.py -i <inputfile> -o <outputfile>')
		sys.exit(2)


if __name__ == "__main__":
	main(sys.argv[1:])
	
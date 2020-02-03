
import os
import json
import AnkiTools
from AnkiTools import anki_convert
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *

class AnkiConverter(QThread):
	def __init__(self, filename):
		super(AnkiTools, self).__init__()
		self.filename = filename

	def import_file(self):
		print(self.filename)


class Fetch(QThread):
	signal = pyqtSignal("PyQt_PyObject")

	def __init__(self, index, id):
		QThread.__init__(self)

		self.index = index
		self.id = id

	def run(self):
		try:
			self.signal.emit(response.json())
		except Exception as e:
			self.signal.emit({
				"data" : {
					'success': False,
					"licenseID" : self.id,
					'message' :{
						"errorText" : "Exceed Max requests {} Retry again".format(repr(e))
					} 
				}
			})

if __name__ == "__main__":
	input_filename = "data/test.apkg"
	output_filename = "data/test.xlsx"
	anki_convert(input_filename, out_file=output_filename)
	# anki_convert('my_workbook.xlsx', out_format='.apkg')

	
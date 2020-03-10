
import json
import AnkiTools
import os
import sys, getopt
from openpyxl import Workbook, load_workbook
import shutil
import xml.sax.saxutils as saxutils
from bs4 import BeautifulSoup as BS, SoupStrainer
import pyexcel_xlsxwx
import pandas as pd
import re
import sqlite3
from zipfile import ZipFile
from filestack import Client


class AnkiConverter():
	def __init__(self, inputfile, filename, filestack_api_key):
		self.inputfile = inputfile
		self.outputfile = filename + ".xlsx"
		self.filename = filename
		self.media_dir = "media"
		self.media_file_path = self.media_dir + "/media"
		self.anki2_file_path = self.media_dir + "/collection.anki2"
		self.filestack_api_key = filestack_api_key
		self.client = Client(self.filestack_api_key)

	def run(self):

		try:
			if self.extract_apkg(self.inputfile):
				self.extract_to_excel(self.outputfile)
			else:
				print("[FAIL]", "Extracting apkg file")
				return
			
		except PermissionError as e:
			print("[FAIL]", "Can not access the file because it is being used by another process {}".format(repr(e)))
			return
		except Exception as e:
			print("[FAIL]", repr(e))
			raise(e) #TODO
		

	def extract_apkg(self, inputfile):

		try:
			#==========Delete media dir if exist=========#
			if os.path.exists(self.media_dir):
				shutil.rmtree(self.media_dir)

			#==========Extract apkg file=================#
			with ZipFile(inputfile, 'r') as zipObj:
				zipObj.extractall(self.media_dir)

			print("[SUCCESS]	Extracting '{}'".format(inputfile))

			with open(self.media_file_path, 'r') as f:
				files = json.load(f)

			self.media_filelist = []
			for key, value in files.items():
				self.media_filelist.append((key, value))
				os.rename("{}/{}".format(self.media_dir, key), "{}/{}".format(self.media_dir, value))
			
			print("[SUCCESS]	Loaded {} media files".format(len(self.media_filelist)))

			return True

		except Exception as e:
			print("[FAIL] ", repr(e))
			return False

	def get_notes_from_anki2(self):

		notes_list = []
		try:
			conn = sqlite3.connect(self.anki2_file_path)
			c = conn.cursor()
			query = """ SELECT id, guid, flds, sfld, mid FROM notes"""
		except Exception as e:
			print("[FAIL]", repr(e))
			raise(e)
			
		try:
			c.execute(query)
			rows = c.fetchall()
			for row in rows:
				data = {}
				data.update({"id" : row[0]})
				data.update({"guid" : row[1]})
				data.update({"html" : row[2]})
				data.update({"text" : row[3]})
				data.update({"mid" : row[4]})
				notes_list.append(data)
		except Exception as e:
			print("[FAIL]", repr(e))
			raise(e)
		c.close()
		conn.close()
		return notes_list

	def get_model_from_anki2(self):
		model = {}
		try:
			conn = sqlite3.connect(self.anki2_file_path)
			c = conn.cursor()
			query = """ SELECT models FROM col"""

		except Exception as e:
			print("[FAIL]", repr(e))
			raise(e)
			
		try:
			c.execute(query)
			rows = c.fetchone()
			model = json.loads(rows[0], encoding="utf-8")
		except Exception as e:
			print("[FAIL]", repr(e))
			raise(e)
		c.close()
		conn.close()
		return model


	def extract_to_excel(self, outputfile):

		#Delete excel file
		if os.path.exists(outputfile):
			os.remove(outputfile)

		note_lists = self.get_notes_from_anki2()
		models = self.get_model_from_anki2()

		anki_list = []
		for note in note_lists:
			anki_item = self.get_anki_item(note, models)
			anki_list.append(anki_item)

		#============SAVE TO EXCEL==============#
		anki_dataframe = pd.DataFrame(anki_list)
		anki_dataframe.to_excel(self.outputfile, sheet_name="Quiz", index=None)
		print('[SUCCESS]', "Successfully saved '{}'".format(self.outputfile))
	

	def get_anki_item(self, note, models):
		
		anki_item = {
			"number" : "",
			"parent_question_number": "",
			"child_question_number": "",
			"subject": "",
			"chapter": "",
			"qtype": "",
			"question" : "",
			"answer1": "",
			"answer2": "",
			"answer3": "",
			"answer4": "",
			"answer5": "",
			"image": "",
			"video": "",
			"info": "",
			"info_image": "",
			"info_video": "",
			"tag1": "",
			"tag2": "",
			"tag3": "",
			"tag4": "",
			"tag5": "",
			"tag6": "",
			"tag7": "",
			"tag8": "",
			"tag9": "",
			"tag10": "",
			"wrong1": "",
			"wrong2": "",
			"wrong3": "",
			"time_limit": "",
			"difficulty": "",
			"is_demo": ""
		}
		note_text = note['html']
		card_id = note['id']
		model_id = str(note['mid'])
		#=======================EXTRACT QUESTION & ANSWER TEXT=================#
		subject = ""
		question = ""
		image = ""
		video = ""
		info_text = ""
		info_image = ""
		info_video = ""
		answers = []

		note_type =  models[model_id].get('type')
		field_count = len(models[model_id].get('flds'))
		
		if note_type == 1:
			question_text = note_text.split("")[0]
			try:
				additional_info = note_text.split("")[1]
			except:
				additional_info = ""
			
			question_answer_text = saxutils.unescape(question_text)
			question_answer_body = BS(question_answer_text, 'lxml').find('body')

			try:
				image = question_answer_body.find('img')['src']
				question_answer_body.img.decompose()
				if image:
					image = self.upload_media_file(image)

			except: image = ""

			try:
				video = question_answer_body.find('video')['src']
				question_answer_body.video.decompose()
				if video:
					video = self.upload_media_file(video)
			except: video = ""
			
			for answer in question_answer_body.find_all(text=re.compile("{{c1::")):
				try:
					# Remove from html tree and get content , error means duplicated
					if "}}" in answer.parent.text:
						answer.parent.replace_with('')
						for item in answer.split("}}"):
							if item.strip():
								#=============Remove {{..}}==========#
								# answers.append(item + "}}")
								answers.append(item.split("c1::")[1])
					else:
						parent = answer.parent
						parent_next = answer.parent.find_next_sibling('div')
						answer_text = answer.parent.text + " " + parent_next.text
						#=============Remove {{..}}==========#
						# answers.append(answer_text)
						answers.append(answer_text.split("c1::")[1].split("}}"))
						parent.replace_with('')
						parent_next.replace_with('')
					
				except Exception as e:
					continue
			

			question_answer_body_childrens = list(question_answer_body.stripped_strings)
			
			try:
				subject = question_answer_body_childrens[0]
			except Exception as e:
				pass
				# print("subject {} :".format(card_id), repr(e), question_answer_body_childrens)

			try:
				question = "".join(question_answer_body_childrens[1:])
			except Exception as e:
				pass
				# print("question {} :".format(card_id), repr(e), question_answer_body)

			try:
				additional_body = BS(saxutils.unescape(additional_info), 'lxml')
				try:
					info_image = additional_body.find('img')['src']
					additional_body.img.decompose()
					if info_image:
						info_image = self.upload_media_file(info_image)
				except:
					info_image = ""

				try:
					info_video = additional_body.find('video')['src']
					additional_body.video.decompose()
					if info_video:
						info_video = self.upload_media_file(info_video)
				except:
					info_video = ""

				info_text = additional_body.text.strip()

			except Exception as e:
				print(repr(e), additional_info)

		else:
			if field_count == 2:

				question_text = note_text.split("")[0]
				try:
					answer_text = note_text.split("")[1]
				except:
					answer_text = ""

				question_answer_body = BS(saxutils.unescape(question_text), 'lxml').find('body')
				answer_body = BS(saxutils.unescape(answer_text), 'lxml').find('body')

				try:
					image = question_answer_body.find('img')['src']
					question_answer_body.img.decompose()
					if image:
						image = self.upload_media_file(image)

				except: image = ""

				try:
					video = question_answer_body.find('video')['src']
					question_answer_body.video.decompose()
					if video:
						video = self.upload_media_file(video)
				except: video = ""
			
				if "___" in question_answer_body.text:

					question_answer_body_childrens = list(question_answer_body.stripped_strings)
					try:
						question = "".join(question_answer_body_childrens)
					except Exception as e:
						print("question {} :".format(card_id), repr(e), question_answer_body)

				else:
					question_answer_body_childrens = list(question_answer_body.stripped_strings)
					try:
						if "?" in question_answer_body_childrens[0]:
							try:
								question = "".join(question_answer_body_childrens)
							except Exception as e:
								print("question {} :".format(card_id), repr(e), question_answer_body)
						else:
							try:
								subject = question_answer_body_childrens[0]
							except Exception as e:
								# print("Subject {} :".format(card_id), repr(e), question_answer_body_childrens)
								pass

							try:
								question = "".join(question_answer_body_childrens[1:])
							except Exception as e:
								print("question {} :".format(card_id), repr(e), question_answer_body)
					except:
						pass
						
				try:
					answers = list(answer_body.stripped_strings)
				except:
					pass

			else:
				pass
				# print(card_id)

		#==============MAKE DATAFRAME============#
		anki_item.update({"number" : card_id})
		anki_item.update({"subject" : subject})
		anki_item.update({"question" : question})

		try: anki_item.update({"answer1" : answers[0]})
		except: anki_item.update({"answer1" : ""})

		try: anki_item.update({"answer2" : answers[1]})
		except: anki_item.update({"answer2" : ""})

		try: anki_item.update({"answer3" : answers[2]})
		except: anki_item.update({"answer3" : ""})

		try: anki_item.update({"answer4" : answers[3]})
		except: anki_item.update({"answer4" : ""})

		try: anki_item.update({"answer5" : answers[4]})
		except: anki_item.update({"answer5" : ""})

		anki_item.update({"image" : image})
		anki_item.update({"video" : video})
		anki_item.update({"info" : info_text})
		anki_item.update({"info_image" : info_image})
		anki_item.update({"info_video" : info_video})

		return anki_item


	def upload_media_file(self, filename):
		try:
			#================TODO WHEN Deploy ===============
			filelink = self.client.upload(filepath="{}/{}".format(self.media_dir, filename))
			print("[SUCCESS] Uploaded Successfully", "'{}'".format(filelink.url))
			return filelink.url
		except Exception as e:
			print("[FAIL] Uploaded Successfully due to", repr(e), filename)
			return None
  
def get_config(filename):
	path = "data/" + filename
	try:
		with open(path, 'r') as f:
			config = json.load(f)
	except Exception as e:
		raise e
	
	return config

def main(argv):
	inputfile  = ''
	
	#==========TEST==========#
	config = get_config("config.json")
	is_test = config.get('DEVELOPMENT')
	
	if is_test:
		print(is_test)
		filestack_api_key = config.get('FILESTACK_TEST_API_KEY')
	else:
		filestack_api_key = config.get('FILESTACK_API_KEY')
		
	try:
		inputfile = argv[1]
		extension = 'apkg'
		filename = inputfile.split(".apkg")[0]
	except:
		print ('AnkiToExcel.py <inputfile.apkg>')
		exit(2)
	
	if extension in ['apkg']:

		ankitoexcel = AnkiConverter(inputfile, filename, filestack_api_key)
		ankitoexcel.run()
	else:
		print ('try: AnkiToExcel.py <inputfile.apkg>')
		exit(2)

if __name__ == "__main__":
	main(sys.argv)
	
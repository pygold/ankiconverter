import requests
import cfscrape as cfs
import json
from csv import DictReader
from log import log
import os
import random
from datetime import datetime
import threading
from time import sleep
from bs4 import BeautifulSoup as BS
from pooky import *
import sys
#Harverster
from recaptchaharvester.Harvester import Harvester
from recaptchaharvester.Fetch import main as getToken

class FileNotFound(Exception):
	''' Raised when a file required for the program to operate is missing. '''

class NoDataLoaded(Exception):
	''' Raised when the file is empty. '''

class OutOfProxies(Exception):
	''' Raised when there are no proxies left '''

class FailCheckoutProcess(Exception):
	''' Raised when checkout process is failed '''

def print_error_log(error_message):
	monitor_time = datetime.now().strftime('%m-%d-%Y %H:%M:%S')
	dir = "./logs/"
	if not os.path.exists(dir):
		os.makedirs(dir)

	with open("logs/error_logs.txt", 'a+') as f:
		f.write("[{}] \t {} \n".format(monitor_time, error_message))


def get_from_csv(filename):
	datas = []
	path = "data/" + filename
	try:
		with open(path, 'r') as f:
			profiles = DictReader(f)
			datas = list(profiles)
	except:
		log('e', "Couldn't locate <" + path + ">.")
		raise FileNotFound()
	
	if (len(datas) == 0):
		raise NoDataLoaded()

	return datas

def get_proxy(proxy_list):
	'''
	(list) -> dict
	Given a proxy list <proxy_list>, a proxy is selected and returned.
	'''
	# Choose a random proxy
	proxy = random.choice(proxy_list)

	# Set up the proxy to be used
	proxies = {
		"http": "http://" + str(proxy),
		"https": "https://" + str(proxy)
	}

	# Return the proxy
	return proxies
 
def get_config(filename):
	path = "data/" + filename
	try:
		with open(path, 'r') as f:
			config = json.load(f)
	except Exception as e:
		raise FileNotFound()
	
	return config


def read_from_txt(filename):
	'''
	(None) -> list of str
	Loads up all sites from the sitelist.txt file in the root directory.
	Returns the sites as a list
	'''
	# Initialize variables
	raw_lines = []
	lines = []

	#Added by ME 
	path = "data/" + filename

	# Load data from the txt file
	try:
		f = open(path, "r")
		raw_lines = f.readlines()
		f.close()

	# Raise an error if the file couldn't be found
	except:
		log('e', "Couldn't locate <" + path + ">.")
		raise FileNotFound()

	# Parse the data
	for line in raw_lines:
		list = line.strip()
		if list != "":
			lines.append(list)

	if(len(lines) == 0):
		raise NoDataLoaded()

	# Return the data
	return lines

def print_error_log(error_message):
	monitor_time = datetime.now().strftime('%m-%d-%Y %H:%M:%S')
	dir = "./logs/"
	if not os.path.exists(dir):
		os.makedirs(dir)

	with open("logs/error_logs.txt", 'a+') as f:
		f.write("[{}] \t {} \n".format(monitor_time, error_message))

def check_keywords_in_product_name(product_name, keywords):
	prod_name = product_name.strip().lower()
	keywords = keywords.replace(' ', '')
	keywords_list = keywords.split(';')

	sum = 0
	p = 0
	for keyword in keywords_list:
		s = keyword[:1]
		k = keyword[1:].lower()
		if s == "+":
			b = 1
			p = p + 1
		elif s == '-':
			b = -1

		if k in prod_name:
			sum = sum + b
	if p == sum:
		return True
	else:
		return False
	

def is_string_exclude(string, substring):
	a = string.string

def send_key_to_licserver(key):
	if key == "123":
		return True
	else:
		return False
	
def check_license_key(config):
	lickey = config.get('key')
	while True:
		if lickey == "":
			log('w', "Please Input Key:")
			lickey = input()
			continue
		elif send_key_to_licserver(lickey):
			log('s', "Valid key!")
			config.update({'key' : lickey})
			with open('data/config.json', 'w', encoding='utf-8') as f:
				json.dump(config, f, ensure_ascii=False)
			break
		else:
			log('f', "Invalid key!")
			log("w", "Please Input Key:")
			lickey = input()
			continue

def main(config, profiles):
	check_license_key(config)
	# returnTime()
	harvester = Harvester("6LeWwRkUAAAAAOBsau7KpuC9AV-6J8mhw4AjC3Xz", "https://www.supremenewyork.com", "127.0.0.1")
	harvester.run()
	ct = CreateTask(profiles=profiles)
	ct.run()

class CreateTask():
	def __init__(self, *args, **kwargs):
		self.profiles = kwargs.get('profiles')

	def run(self):
		log('w', "Welcome SupremeBot 1.11")
		proxies_uk = read_from_txt('proxies_uk.txt')
		proxies_us = read_from_txt('proxies_us.txt')

		for profile in self.profiles:
			if profile.get('country').lower() == 'gb':
				proxies = proxies_uk
			elif profile.get('country').lower() == 'us':
				proxies = proxies_us

			sb = SupremeBot(profile=profile, proxies=proxies)
			sb.start()

class SupremeBot(threading.Thread):
	def __init__(self, *args, **kwargs):
		threading.Thread.__init__(self)
		self.profile = kwargs.get('profile')
		self.proxies = kwargs.get('proxies')
		self.shop_url   = "https://www.supremenewyork.com/shop"
		self.base_url   = "https://www.supremenewyork.com"
		self.user_agent         = "Mozilla/5.0 (Windows NT 6.2; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.78 Safari/537.36"
		self.session = requests.session()
		self.scraper = cfs.create_scraper(delay=10, sess=self.session)
		self.product_to_buy = []

	def find_product_using_keywords(self):
		headers = {
			"User-Agent": self.user_agent,
			"Upgrade-Insecure-Requests": "1",
			"Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3",
		}
		while True:
			self.proxy = get_proxy(self.proxies)
			try:
				response = self.scraper.get(self.shop_url, timeout=5, proxies=self.proxy, headers=headers)
			except Exception as e:
				log('f', "Connection Refused due to proxy error, Changing proxy")
				log('f', str(e))
				continue
			if response.status_code == 200:
				prod_html = BS(response.content, features='lxml').find(id="shop-scroller")
				category_name = self.profile.get('category').lower()
				for item in prod_html.find_all("li", category_name):
					prod_url = self.base_url + item.find('a').get('href')
					try:
						self.find_product_using_color(prod_url)
					except Exception as e:
						log('f', str(e))
				break
			elif response.status_code == 403:
				log('f', "{} Access Denied! Changing Proxy".format(str(response.status_code)))
				continue
			else:
				log('f', "{} Status code! Changing Proxy".format(str(response.status_code)))
				continue
				

	def find_product_using_color(self, url):
		headers = {
			"User-Agent": self.user_agent
		}
		try:
			response = self.scraper.get(url, timeout=5, proxies=self.proxy, headers=headers)
		except Exception as e:
			log('f', str(e))

		if response.status_code == 200:
			prod_name   = BS(response.content).find('h1', 'protect').string
			keywords    = self.profile['keywords']
			if check_keywords_in_product_name(prod_name, keywords):
				for li in BS(response.content).find(id="details").find('ul').find_all('li'):
					color = li.find('a').get('data-style-name')
					is_sold_out = li.find('a').get('data-sold-out')
					if self.profile.get('colour').lower() == color.lower() and is_sold_out == "false":
						prod_style_url = self.base_url + li.find('a').get('href')
						self.find_product_using_size(prod_style_url)
					
		elif response.status_code == 403:
			log('f', "Fail Get Product Detail Due to Access Denied 403")
		else:
			log('f', "Fail Get Product Detail: {}".format(str(response.status_code)))


	def find_product_using_size(self, url):
		try:
			response = self.scraper.get(url, timeout=5, proxies=self.proxy)
		except Exception as e:
			log('f', "Connection Refused due to proxy connection failure: {}".format(str(e)))

		if response.status_code == 200:
			try:
				form_block = BS(response.content, features='lxml').find(id="cart-addf")
				action_url = self.base_url + form_block.get('action')
				style_id   = form_block.find(id="style").get('value')
				for size in form_block.find_all('option'):
					if size.string.lower() == self.profile.get('size').lower():
						size_value = size.get('value')

				if size_value:
					log('s', "Found product!")
					product = {}
					product.update({'action_url': action_url})
					product.update({'style_id': style_id})
					product.update({'size_value': size_value})
					product.update({'commit' : 'add to basket'})
					product.update({'utf-8': '✓'})
					self.add_to_cart(product)
				else:
					log('f', "Can not find product matching with size {}".format(self.profile.get('size')))
					
			except Exception as e:
				raise(e)
		elif response.status_code == 403:
			log('f', "{} Access Denied!".format(str(response.status_code)))
		else:
			log('f', "{} Status code!".format(str(response.status_code)))

	def add_to_cart(self, product):
		headers = {
			"User-Agent": self.user_agent,
			"Upgrade-Insecure-Requests": "1",
			"Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3",
			"X-Requested-With": "XMLHttpRequest"
		}
		atc_url = product.get('action_url')
		data = {
			"utf8": product.get('utf-8'),
			"style": product.get('style_id'),
			"size" : product.get('size_value'),
			"commit": product.get("commit")
		}
		response = self.scraper.post(atc_url, data=data, proxies=self.proxy, timeout=5, headers=headers)

		if response.status_code == 200:
			log('s', "ADD TO CART Success!")
			self.make_billing()
		else:
			log('f', response.status_code)

	def anti_recaptcha(self):
		pass

	def make_billing(self):
		headers = {
			"User-Agent": self.user_agent,
			"Upgrade-Insecure-Requests": "1",
			"Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3",
		}
		checkout_url = "https://www.supremenewyork.com/checkout"
		try:
			response = self.scraper.get(checkout_url, headers=headers, proxies=self.proxy, timeout=5)
			print(response.headers)
			print(response.request.headers)
			print(self.scraper.cookies)
		except Exception as e:
			log('f', str(e))
		
		if response.status_code == 200:
			# with open('dummy/checkout.html', 'wb') as f:
			#     f.write(response.content)
			checkout_html = BS(response.content, features='lxml').find(id="checkout_form")
			checkout_payload = {}
			for item in checkout_html.find_all('input'):
				checkout_payload.update({item.get('name'): item.get('value')})

			checkout_payload.update({'order[billing_name]': self.profile.get('fname') + " " + self.profile.get('lname')})
			checkout_payload.update({'order[email]': self.profile.get('email')})
			checkout_payload.update({'order[tel]': self.profile.get('phone')})
			checkout_payload.update({'order[billing_address]': self.profile.get('line1')})
			checkout_payload.update({'order[billing_address_2]': self.profile.get('housenum')})
			checkout_payload.update({'order[billing_address_3]': ""})
			checkout_payload.update({'order[billing_city]': self.profile.get('city')})
			checkout_payload.update({'order[billing_zip]': self.profile.get('postcode')})
			checkout_payload.update({'order[billing_country]': self.profile.get('country')})
			checkout_payload.update({'same_as_billing_address': '1'})
			checkout_payload.update({'store_credit_id': ''})
			checkout_payload.update({'store_address': '1'})
			checkout_payload.update({'credit_card[type]': 'visa'})
			checkout_payload.update({'credit_card[cnb]': self.profile.get('number')})
			checkout_payload.update({'credit_card[month]': self.profile.get('expirymonth')})
			checkout_payload.update({'credit_card[year]': self.profile.get('expiryyear')})
			checkout_payload.update({'credit_card[ovv]': self.profile.get('cvc')})
			checkout_payload.update({'hpcvv': ''})
			##TODO
			checkout_payload.update({'g-recaptcha-response': ''})
			checkout_payload.update({'cardinal_id': ''})
			checkout_payload.update({'cardinal_jwt': ''})
			self.product_check_out(checkout_payload)
		else:
			log('f', str(response.status_code))

	def product_check_out(self, checkout_payload):
		print(checkout_payload)


	def run(self):
		self.find_product_using_keywords()


def returnTime():

	timeS = "Please enter the hour of the drop (e.g. if you're in the UK enter '11' for 11.00am) "
	   
	isTime = input(timeS)
	try:
		dropTime = int(isTime)
	except ValueError:
		log('f', "You must input number!")
		exit()
	text = "Drop selected for "
	text += isTime
	text += ":00"
	text += " the program will automatically run at this time. Please do not close cmd box."
	print(text)
	while True:
		ts = time.time()
		houra = datetime.fromtimestamp(ts).strftime('%H')
		hour = int(houra)
		old_time = ""
		cur_time = datetime.fromtimestamp(ts).strftime('%H:%M:%S')
		if cur_time != old_time:
			sys.stdout.write('\r'+ cur_time)
			old_time = datetime.fromtimestamp(ts).strftime('%H:%M:%S')
		if hour >= dropTime:
			break

if (__name__=="__main__"):
	try:
		config = get_config("config.json")
	except Exception as e:
		log('f', "Load Config File")
		print_error_log("Fail Load Config file:")
		raise(e)
	try:
		profiles = get_from_csv("profiles.csv")
	except Exception as e:
		print_error_log("Fail Load Profiles file:")
		raise(e)
	main(config, profiles)


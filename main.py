import requests
# from openpyxl import Workbook
# from openpyxl import load_workbook
# from bs4 import BeautifulSoup

from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
# import pyautogui	#pip install PyAutoGUI
from webdriver_manager.chrome import ChromeDriverManager

url = f"https://www.whtop.com/directory/category/shared-hosting"

options = Options()
user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36'
options.add_argument('user-agent={0}'.format(user_agent))
accept_language = "en-US,en;q=0.9"
options.add_argument('accept-language={0}'.format(accept_language))
# options.add_argument("user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.163 Safari/537.36")
# options = webdriver.ChromeOptions()
#  driver = webdriver.Chrome(executable_path="D:/My_PyCharm/Programs for python/Chromedriver/chromedriver.exe",
# 						  options=options)
driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
wait = WebDriverWait(driver, 20)
hhh = driver.get(url)
time.sleep(5)
source_html = driver.page_source
driver.close()
# запис в ХТМЛ-файл
with open(f"index1.html", 'w', encoding='utf-8') as file:
	file.write(source_html)

import lxml, re
from bs4 import BeautifulSoup
import json
#
# links_result = []
# HTMLFile = open(f"4660 providers of shared_budget web hosting. Directory.html", "r", encoding='utf-8')
# # Reading the file
# index = HTMLFile.read()
# soup = BeautifulSoup(index, 'lxml')
# # 5 пошук потрібних значень
# # url_fin = soup.find('div', id='les_membres_du_reseau')
# url_next = soup.find_all('div', class_='company clearer')
# # print(url_next)
# for item_list in url_next:
# 	find_title = item_list.find('div', class_='company-title')
# 	name = find_title.find('b').find('a').text
# 	# print(name)
# 	url_title = find_title.find('small').text
# 	# print(url_title)
# 	find_discription = item_list.find('div', class_='company-description aj')
# 	location = find_discription.find('span', class_='gray').text
# 	# print(location)
# 	description = find_discription.find('span',property='description').text
# 	find_rank = item_list.find('ul', class_='alphabetical').find_all('span', class_='rank')
# 	# rank_name = find_rank[row].find('span').text.strip()
# 	Alexa_rating = find_rank[0].find('span',class_='value').text
# 	MOZ_DA_PA = find_rank[1].find('span',class_='value').text
# 	Links_Count = find_rank[2].find('span',class_='value').text
# 	links_result.append(
# 		{
# 			'url': url_title,
# 			'name': name,
# 			'location': location,
# 			'description': description,
# 			'Alexa_rating': Alexa_rating,
# 			'MOZ_DA_PA': MOZ_DA_PA,
# 			'Links_Count': Links_Count
# 		}
# 	)
#
# # 	6 запис у файл JSON
# with open(f"result.json", 'w', encoding='utf-8') as file:
# 	json.dump(links_result, file, indent=4, ensure_ascii=False)

# # 6 Запис в ексель
# import xlsxwriter #pip install XlsxWriter
# import os
#
# # workbook = xlsxwriter.Workbook(f"NBU{dt.date.today().strftime('%d%m%Y')}.xlsx")
# workbook = xlsxwriter.Workbook("D:\Path\Vhtop_com.xlsx")
#
# bold = workbook.add_format({'bold': True, 'font_color': 'red'})
# bold.set_align('center')
#
# bold_1 = workbook.add_format({'bold': True, 'font_color': 'black'})
# bold_1.set_align('center')
# bold_1.set_border(1)
#
# bold_2 = workbook.add_format({'bold': True, 'font_color': 'blue'})
# bold_2.set_align('center')
# bold_2.set_border(1)
#
# bold_3 = workbook.add_format({'bold': True, 'font_color': 'black'})
# # bold_3.set_bg_color('#b4b4b4')
# bold_3.set_align('center')
# bold_3.set_border(2)
#
# dirPath = r"D:\Path"
# extensions = (".json")
#
# for root, dirs, files in os.walk(dirPath):
# 	for filename in files:
# 		if os.path.isfile(os.path.join(dirPath, filename)):
# 			if filename.endswith(extensions):
# 				file_name = os.path.join(dirPath, filename)
# 				worksheet = workbook.add_worksheet(name=f"{str(filename).split('.')[0]}")
# 				# Format the column
# 				worksheet.set_column('A:A', 25)
# 				worksheet.set_column('B:B', 20)
# 				worksheet.set_column('C:C', 25)
# 				worksheet.set_column('D:D', 35)
# 				worksheet.set_column('E:E', 15)
# 				worksheet.set_column('F:F', 15)
# 				worksheet.set_column('G:G', 15)
# 				# worksheet.set_column('H:H', 10)
# 				worksheet.set_default_row(25)
#
# 				worksheet.write('A1', 'Host url', bold_3)
# 				worksheet.write('B1', 'name', bold_3)
# 				worksheet.write('C1', 'Location', bold_3)
# 				worksheet.write('D1', 'Company description', bold_3)
# 				worksheet.write('E1', 'Alexa rating', bold_3)
# 				worksheet.write('F1', 'MOZ DA/DP', bold_3)
# 				worksheet.write('G1', 'Links Count', bold_3)
# 				# worksheet.write('H1', 'State', bold_3)
# 				file_json = open(file_name, 'r')
# 				row = 2
# 				data = json.load(file_json)
#
# 				for i in data:
# 					# add new row
# 					worksheet.write(f'A{row}', i['url'])
# 					worksheet.write(f'B{row}', i['name'])
# 					worksheet.write(f'C{row}', i['location'])
# 					worksheet.write(f'D{row}', i['description'])
# 					worksheet.write(f'E{row}', i['Alexa_rating'])
# 					worksheet.write(f'F{row}', i['MOZ_DA_PA'])
# 					worksheet.write(f'G{row}', i['Links_Count'])
# 					# worksheet.write(f'H{row}', i['state'])
# # 					worksheet.write(f'B{row}', valute, bold_1)
# # 					worksheet.write(f'C{row}', rate, bold_1)
# 					row += 1
# time.sleep(4)
# workbook.close()
# print(f"Записано файл: {workbook.filename}")
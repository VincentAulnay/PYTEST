print('start')
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import xlwt
import xlrd
import time
from xlutils.copy import copy
import datetime
import datetime as dt
from tkinter import filedialog
from tkinter import *
from bs4 import BeautifulSoup
from urllib.request import urlopen
import os
import shutil
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from openpyxl import load_workbook
import threading
import sys



chrome_options = webdriver.ChromeOptions()
#prefs = {"profile.managed_default_content_settings.images": 2}
#chrome_options.add_experimental_option("prefs", prefs)
chrome_options.add_argument("window-size=2000,1000")
#chrome_options.add_argument("-headless")
#chrome_options.add_argument("-disable-gpu")
print ('▀▄▀▄▀▄ STOPBNB ▄▀▄▀▄▀')

now = str(datetime.datetime.now())[:19]
now = now.replace(":","_")
print(now)
print('tes')
#-----EXCEL RESULT OPEN AND READ-----

import re
import json
import csv
from google.oauth2 import service_account
import pygsheets
import pandas as pd

#wbx = load_workbook(path_RESULT.filename)
#ws = wbx.active
print('ici1')
try:
	client = pygsheets.authorize(service_account_file='/home/vincent/Desktop/raspbian-364809-be26e1ee6573.json')
except:
	client = pygsheets.authorize(service_account_file='/home/pi/Desktop/raspbian-364809-be26e1ee6573.json')
print('ici1')
#spreadsheet_url = "https://docs.google.com/spreadsheets/d/1vx34zctZXc2eQSFFe4I7zY6bjKJz9MtO7pgAIaQix4c/edit?usp=sharing"
spreadsheet_url = "https://docs.google.com/spreadsheets/d/1ACSlRUHdqn9ExIM2M-18VGHBoo8RaxAfxkbvKCd1ylw/edit?usp=sharing"
print('ici2')	

sheet_data = client.sheet.get('1ACSlRUHdqn9ExIM2M-18VGHBoo8RaxAfxkbvKCd1ylw')
print('ici3')
sheet = client.open('ZG')
print('ici4')
ws = sheet.worksheet_by_title('Sheet1')
print('sheet')
#print(ws)
#-------FIND COLUMN UPDATE------
cTITLE=2
cANNONCE=3
cNAME_HOTE=4
cANCIENNETE=5
cHOTE=6
cPRICE=7
cCOMMENT=8
cTYPE_LOGEMENT=9
cVOYAGEUR=10
cCHAMBRE=11
cSdB=12
cLITS=13
cVILLE=14
clat=15
clon=16
cSUPERHOTE=17
cCOMMENT_PROFIL=18
cID_VERIF=19
cNB_ANNONCE=20
cPROPRETE=21
cPRECISION=22
cCOMMUNICATION=23
cEMPLACEMENT=24
cARRIVEE=25
cQUALITY_PRICE=26
cREGISTER=27
cTAUX_REPONSE=28
cDELAI_REPONSE=29
cCHECK_IN=30
cCHECK_OUT=31
cFUMEUR=32
cENFANT=33
cSERRURE=34
cANIMAUX=35
cCAUTION=36
cFUMEE=37
cMONOXYDE=38
cDISTANCIATION_SOCIAL=39
cLANGUE=40
cIMAGE_PROFIL=41
cIMAGE_1=42
cIMAGE_2=43
cIMAGE_3=44
cIMAGE_4=45
cIMAGE_5=46
cCOHOTE_URL1=47
cCOHOTE_NAME1=48
cCOHOTE_IMAGE1=49
cNB_COHOTE=50
cCOHOTE_URL2=51
cCOHOTE_NAME2=52
cCOHOTE_IMAGE2=53
cFETE=54
cACTIVE=55
cENTREPRISE=56
cSOCIALWASHING=57


		
#driver = webdriver.Chrome('/usr/lib/chromium-browser/chromedriver',chrome_options=chrome_options)
driver = webdriver.Chrome('/usr/lib/chromium-browser/chromedriver',chrome_options=chrome_options)
driver.set_window_size(2000, 1000)
driver.get('https://www.google.com/')
driver.get('chrome://settings/')
driver.execute_script('chrome.settingsPrivate.setDefaultZoom(0.5);')
driver.implicitly_wait(10)
wait = WebDriverWait(driver, 5)

#c = ligne 2 du xls resultant
c=2
wait2 = WebDriverWait(driver, 5)
wait3 = WebDriverWait(driver, 5)
wait = WebDriverWait(driver, 2)

def scrap(h):
	global scrap_ok
	driver.get(h)
	time.sleep(2)
	scrap_ok=1

def GSwrite(c, clevel, valeur):
	ws.cell((c, clevel)).value = valeur
	
#threading.Thread(target=GSwrite, args=(c, clevel, valeur,)).start()
	
fm=2
fff=0

nrow=100000

while c<=nrow:
	scrap_ok=0
	print (str(c)+'/'+str(nrow))
	h=ws.cell((c, cANNONCE)).value
	numero=None
	#print (h)
	#do=sheet_read.cell(i,0).value
	if numero is None:
		threading.Thread(target=scrap, args=(h,)).start()
		timer=1
		while timer<=60:
			if scrap_ok==1:
				try:
					f_ele=0
					while f_ele<=3:
						try:
							button_fermer = wait.until(EC.presence_of_element_located((By.XPATH, "//button[@aria-label='Fermer']")))
							button_fermer.click()
						except:
							pass
						try:
							#ele=driver.find_element_by_xpath("//div[@class='_1cvivhm']")
							#ele=driver.find_element_by_xpath("//div[@class='_cg8a3u']")
							driver.execute_script("window.scrollBy(0,2000);")
							ele=driver.find_element_by_xpath("//div[@class='s9fngse dir dir-ltr']")
							driver.execute_script("arguments[0].scrollIntoView(true);", ele)
							#driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
							time.sleep(1)
							driver.execute_script("window.scrollBy(0,-400);")
							#driver.execute_script("window.scrollBy(0,500);")
							f_ele=6
						except:
							#driver.execute_script("window.scrollBy(0,1000);")
							f_ele=f_ele+1
							time.sleep(2)
					#driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
				#PROFILE
					time.sleep(1)
					html = driver.page_source
					soup = BeautifulSoup(html, 'html.parser')
					try:
						#GPS
						try:
							tp_c=soup.find('div', attrs={"class": "gm-style"})
							tt=tp_c.find('div', attrs={"style": "margin: 0px 5px; z-index: 1000000; position: absolute; left: 0px; bottom: 0px;"})
							tt1=tt.find('a', attrs={"target": "_blank"})
							tt2=tt1['href']
							#----------Create translation table----------
							table = str.maketrans('=&', '++')
							result_gps = tt2.translate(table)
							split_gps=result_gps.split("+")
							#https://www.google.com/maps?ll+46.23657,3.92341+z=14&t=m&hl=fr&gl=FR&mapclient=apiv3
							#https://maps.google.com/maps?ll=49.19054,-2.11715&z=14&t=m&hl=fr&gl=FR&mapclient=apiv3
							#https://www.google.com/maps?ll=48.6472,-2.0054&z=14&t=m&hl=fr&gl=FR&mapclient=apiv3
							coor=split_gps[1]
							long_lat=coor.split(',')
							#--------------Write results--------------
							#print(long_lat[0])
							#print(long_lat[1])
							#point=Point(float(long_lat[1]),float(long_lat[0]))
							#question=polygon.contains(point)
							question=True
							#print(question)
							#sheet.write(c, 12, long_lat[0])
							#sheet.write(c, 13, long_lat[1])
							#ws.cell(row=c+1, column=13+1).value = long_lat[0]
							#ws.cell(row=c+1, column=14+1).value = long_lat[1]
						except:
							question=False
							print('NO GPS')
						if question is True:
							#sheet.write(c, 12, long_lat[0])
							#sheet.write(c, 13, long_lat[1])
							try:
								#ws.cell((c, cANNONCE)).value = h
								#ws.cell((c, clat)).value = long_lat[0]
								threading.Thread(target=GSwrite, args=(c, clat, long_lat[0],)).start()
								#ws.cell((c, clon)).value = long_lat[1]
								threading.Thread(target=GSwrite, args=(c, clon, long_lat[0],)).start()
							except:
								ee=1
						#TITLE
							try:
								div1=soup.find('h1', attrs={"class": "_fecoyn4"})
								#ws.cell((c, cTITLE)).value = div1.text
								threading.Thread(target=GSwrite, args=(c, cTITLE, div1.text,)).start()
								#print(div1.text)
							except:
								#print('NO TITLE')
								aaa=1
						#URL HOTE
							try:
								div=soup.find('div', attrs={"class": "c6y5den dir dir-ltr"})
								div2=div.find('a')
								div1=div2['href']  #.attrs['href']
								ws.cell((c, cHOTE)).value = "https://www.airbnb.fr"+str(div1)
								print("URLHOT1"+str(div1))
							except:
								try:
									div=soup.find('div', attrs={"class": "_dbynel"})
									div2=div.find('a')
									div1=div2['href']  #.attrs['href']
									ws.cell((c, cHOTE)).value = "https://www.airbnb.fr"+str(div1)
									#print("URLHOT2"+str(div1))
								except:
									#print('NO PROFILE')
									aaa=1
						#COMMENTAIRE
							COMMENT='NO COMMENT'
							#run_price=extract("//span[@class='_wfad8t']",6,COMMENT,c,YN_comment)
							try:
								p_c=[]
								try:
									tp_c=soup.find('span', attrs={"class": "_142pbzop"}).text
									print(tp_c)
									#print('type1')

								except:
									try:
										tp_c=soup.findAll('span', attrs={"class": "_2qpirtt"})[1].text
										print(tp_c)
									except:
										try:
											tp_c=soup.find('span', attrs={"class": "_1qx9l5ba"}).text
											print(tp_c)
										except:
											try:
												tp_c=soup.findAll('span', attrs={"class": "_bq6krt"})[1].text
												print(tp_c)
											except:
												#print('type2')
												aaa=1
								p_c=tp_c.replace("(","")
								cc=p_c.replace(")","")
								try:
									pp=cc.split(' ')
									cc=pp[0]
									ws.cell((c, cCOMMENT)).value = cc
								except:
									pass
								ws.cell((c, cCOMMENT)).value = cc
								#print ("COMMENT ===")
								#print(cc)
								#p_c=tp_c.split("(")
								#print('ici1')
								#table_c = p_c[1].replace(")"," ")
								#print (table_c)
								#ws.cell(row=c+1, column=6+1).value = table_c
							except:
								#print('NOCOMMENT')
								aaa=1
						#VOYAGEUR
							try:
								the_tr= soup.find('div', attrs = {'class' : '_tqmy57'})
								tt=the_tr.find_all('li')[0]
								tt1=tt.find_all('span')[0]
								tt2=tt1.text
								p_tp=tt2.split(" ")
								ws.cell((c, cVOYAGEUR)).value = p_tp[0]
								#print("V="+str(p_tp[0]))
							except:
								#print('NO VOYAGER')
								aaa=1

						#LITS
							try:
								the_tr= soup.find('div', attrs = {'class' : '_tqmy57'})
								tt=the_tr.find_all('li')[2]
								tt1=tt.find_all('span')[2]
								tt2=tt1.text
								p_tp=tt2.split(" ")
								ws.cell((c, cLITS)).value = p_tp[0]
								print("L="+str(p_tp[0]))
							except:
								#print('NO LIT')
								aaa=1
						#SdB
							try:
								the_tr= soup.find('div', attrs = {'class' : '_tqmy57'})
								tt=the_tr.find_all('li')[3]
								tt1=tt.find_all('span')[-1]
								tt2=tt1.text
								p_tp=tt2.split(" ")
								ws.cell((c, cSdB)).value = p_tp[0]
								#print("B="+str(p_tp[0]))
							except:
								#print('NO SdB')
								aaa=1
						#CHAMBRE
							try:
								the_tr= soup.find('div', attrs = {'class' : '_tqmy57'})
								tt=the_tr.find_all('li')[1]
								tt1=tt.find_all('span')[2]
								tt2=tt1.text
								p_tp=tt2.split(" ")
								ws.cell((c, cCHAMBRE)).value = p_tp[0]
								print("C="+str(p_tp[0]))
							except:
								#print('NO CHAMBRE')
								aaa=1
						#VILLE
							try:
								tt=soup.find('span', attrs={"class": "_pbq7fmm"}).text
								ws.cell((c, cVILLE)).value = tt
								print(tt)
							except:
								try:
									tt=soup.find_all('a', attrs={"class": "_5twioja"})[-1].text
									ws.cell((c, cVILLE)).value = tt
									print(tt)
								except:
									try:
										tp_c=soup.find('div', attrs={"class": "_9ns6hl"})
										tt=tp_c.find('h3').text
										ws.cell((c, cVILLE)).value = tt
										print(tt)
									except:
										try:
											tp_c=soup.find('span', attrs={"class": "_8vvkqm3"}).text
											ws.cell((c, cVILLE)).value = tp_c
											print(tp_c)
										except:
											try:
												tp_c=soup.find('span', attrs={"class": "_9xiloll"}).text
												ws.cell((c, cVILLE)).value = tp_c
												print(tp_c)
											except:
												print('NO VILLE')
												aaa=1

							#print(tt)
						#NAME_HOTE
							try:
								tp_c=soup.find_all('div', attrs={"class": "hnwb2pb dir dir-ltr"})[0].text
								#print("schema 2")
								pp=tt_c.split('par ')
								ws.cell((c, cNAME_HOTE)).value = pp[1]
								#print(pp[1])
							except:
								try:
									tp_c=soup.find('div', attrs={"class": "_f47qa6"})
									tt=tp_c.find('div', attrs={"class": "_svr7sj"})
									tt1=tt.h2.get_text()
									pp=tt1.split('par ')
									ws.cell((c, cNAME_HOTE)).value = pp[1]
									#print(pp[1])
								except:
									try:
										#tp_c=soup.find_all('h2', attrs={"class": "hnwb2pb dir dir-ltr"})[1].text
										tp_c= soup.find('h2', text=re.compile(r"\bProposé\b")).text
										#print("schema 3")
										pp=tp_c.split('par ')
										ws.cell((c, cNAME_HOTE)).value = pp[1]
										#print(pp[1])
									except:
										aaa=1
								#print ('NO_NAME')
						#TYPE_HOME
							try:
								the_tr= soup.find('div', attrs={"class": "_cv5qq4"})
								ttt=the_tr.h2.text
								pp=ttt.split('⸱')
								print(pp)
								ws.cell((c, cTYPE_LOGEMENT)).value = pp[0]
							except:
								try:
									the_tr= soup.find('div', attrs={"class": "_xcsyj0"}).text
									pp=the_tr.split('.')
									ws.cell((c, cTYPE_LOGEMENT)).value = pp[0]
								except:
									#print('NOTYPE')
									aaa=1

						#ANCIENNETE
							try:
								tp_c=soup.find('div', attrs={"class": "s9fngse dir dir-ltr"}).text
								pp=tp_c.split("depuis")
								ws.cell((c, cANCIENNETE)).value = pp[1]
								#print(tp_c)
							except:
								try:
									tp_c=soup.find('div', attrs={"class": "_f47qa6"})
									tt=tp_c.find('div', attrs={"class": "_svr7sj"})
									tt1=tt.div.get_text()
									ws.cell((c, cANCIENNETE)).value = tt1
								except:
									aaa=1
						#SUPER HOTE
							try:
								#the_tr= soup.find('span', text=re.compile(r'\bSuperhost\b'),attrs = {'aria-hidden' : 'false'})
								#tp_c=soup.find('span', attrs={"class": "_63km3vu"}, text=re.compile(r'\bSuper\b'))
								tp_c=soup.find('span', attrs={"class": "_1mhorg9"})
								if tp_c is not None:
									ws.cell((c, cSUPERHOTE)).value = 'X'
									print (tp_c)
							except:
								aaa=1
						#COMMENT PROFIL
							try:
								the_tr= soup.findAll('li', attrs = {'class' : '_1belslp'})[0]
								the_li= the_tr.find('span', attrs = {'class' : '_pog3hg'})
								ccc=the_li.text
								pp=ccc.split('c')
								cc=pp[0]
								#div2=the_tr.findNextSibling('div')
								#print(the_tr.section.span.div.span.text)
								if cc=='Identité vérifiée':
									cc=0
								ws.cell((c, cCOMMENT_PROFIL)).value = cc
								#print(div2.text)
							except:
								try:
									tp_c=soup.find('ul', attrs={"class": "tq6hspd h1aqtv1m dir dir-ltr"})
									sp=tp_c.find('span', text=re.compile(r'\bcommentaires\b')).text
									#sp=tp_c.findAll('span', attrs={"class": "l1dfad8f dir dir-ltr"})[0]
									pp=sp.split('c')
									cc=pp[0]
									if cc=='Identité vérifiée':
										cc=0
									ws.cell((c, cCOMMENT_PROFIL)).value = cc
								except:
									#print('No Comment profil')
									aaa=1
						#IDENTIFIE CHECK
							try:
								the_tr= soup.find('span', text=re.compile(r"\bIdentité vérifiée\b"))
								ws.cell((c, cID_VERIF)).value = 'YES'
								#print(div2.text)
							except:
								ws.cell((c, cID_VERIF)).value = 'NO'
								#print('No CHECK ID')
								aaa=1
						#CO HOTE
							ifcohote=0

							try:
								the_tr33= soup.find('ul', attrs = {'class' : '_kaabnn'})
								#print(the_tr33)
								ifcohote=3
								#print("cohote3")
								if the_tr33 is None:
									print(bbb)
							except:
								try:
									the_tr22= soup.find('ul', attrs = {'class' : 'ato18ul dir dir-ltr'})
									ifcohote=2
									#print("cohote2")
									if the_tr22 is None:
										print(bbb)
								except:
									try:
										the_tr11= soup.find('ul', attrs = {'class' : '_1omtyzc'})
										ifcohote=1
										#print("cohote1")
									except:
										aaa=1
							if ifcohote==1:
								try:
									the_tr1= the_tr11.findAll('li', attrs = {'class' : '_108byt5'})[0]
									tt11= the_tr1.find('a', attrs = {'target' : '_blank'})
									div1=tt11['href']  #.attrs['href']
									#print(div1)
									ws.cell((c, cCOHOTE_URL1)).value = "https://www.airbnb.fr"+str(div1)
									tt12= the_tr1.find('span', attrs = {'class' : '_1kfl0pr'})
									ws.cell((c, cCOHOTE_NAME1)).value = tt12.text
									#print(tt12.text)
									try:
										tt13= the_tr1.find('img', attrs = {'class' : '_9ofhsl'})
										ws.cell((c, cCOHOTE_IMAGE1)).value = tt13['src']
										#print(tt13['src'])
									except:
										noimage=1
									#2COHOTE
									the_tr2= the_tr11.findAll('li', attrs = {'class' : '_108byt5'})[1]
									tt21= the_tr2.find('a', attrs = {'target' : '_blank'})
									div2=tt21['href']  #.attrs['href']
									#print(div2)
									ws.cell((c, cCOHOTE_URL2)).value = "https://www.airbnb.fr"+str(div2)
									ws.cell((c, cNB_COHOTE)).value = 2
									tt22= the_tr2.find('span', attrs = {'class' : '_1kfl0pr'})
									ws.cell((c, cCOHOTE_NAME2)).value = tt22.text
									#print(tt22.text)
									tt23= the_tr2.find('img', attrs = {'class' : '_9ofhsl'})
									ws.cell((c, cCOHOTE_IMAGE2)).value = tt23['src']
									#print(tt23['src'])
								except:
									try:
										the_tr1= the_tr11.find('li', attrs = {'class' : '_108byt5'})
										tt= the_tr1.find('a')
										div1=tt['href']  #.attrs['href']
										#print(div1)
										ws.cell((c, cCOHOTE_URL1)).value = "https://www.airbnb.fr"+str(div1)
										tt2= the_tr1.find('span')
										ws.cell((c, cCOHOTE_NAME1)).value = tt2.text
										#print(tt2.text)
										tt3= the_tr1.find('img', attrs = {'class' : '_6tbg2q'})
										ws.cell((c, cCOHOTE_IMAGE1)).value = tt3['src']
										#print(tt3['src'])
										ws.cell((c, cNB_COHOTE)).value = 1
									except:
										#ws.cell((c, cNB_COHOTE).value = 0
										#print('no co hote')
										aaa=1
							if ifcohote==2:
								#print("start cohote 2")
								try:
									the_tr1= the_tr22.findAll('li', attrs = {'class' : 'ahxgcvj dir dir-ltr'})[0]
									#print("li2 find")
									#print(the_tr1)
									tt11= the_tr1.find('a')
									#print("a find")
									div1=tt11['href']  #.attrs['href']
									ws.cell((c, cCOHOTE_URL1)).value = "https://www.airbnb.fr"+str(div1)
									tt12= the_tr1.find('span')
									ws.cell((c, cCOHOTE_NAME1)).value = tt12.text
									try:
										tt13= the_tr1.find('img')
										ws.cell((c, cCOHOTE_IMAGE1)).value = tt13['src']
									except:
										noimage=1
									#2COHOTE
									the_tr2= the_tr22.findAll('li', attrs = {'class' : 'ahxgcvj dir dir-ltr'})[1]
									tt21= the_tr2.find('a')
									div2=tt21['href']  #.attrs['href']
									ws.cell((c, cCOHOTE_URL2)).value = "https://www.airbnb.fr"+str(div2)
									ws.cell((c, cNB_COHOTE)).value = 2
									tt22= the_tr2.find('span')
									ws.cell((c, cCOHOTE_NAME2)).value = tt22.text
									tt23= the_tr2.find('img')
									ws.cell((c, cCOHOTE_IMAGE2)).value = tt23['src']
								except:
									try:
										the_tr1= the_tr22.find('li', attrs = {'class' : 'ahxgcvj dir dir-ltr'})
										#print("li1 find")
										tt11= the_tr1.find('a')
										div1=tt11['href']  #.attrs['href']
										ws.cell((c, cCOHOTE_URL1)).value = "https://www.airbnb.fr"+str(div1)
										tt12= the_tr1.find('span')
										ws.cell((c, cCOHOTE_NAME1)).value = tt12.text
										tt13= the_tr1.find('img')
										ws.cell((c, cCOHOTE_IMAGE1)).value = tt13['src']
										ws.cell((c, cNB_COHOTE)).value = 1
									except:
										aaa=1
							if ifcohote==3:
								try:
									the_tr1= the_tr33.findAll('li', attrs = {'class' : '_108byt5'})[0]
									tt11= the_tr1.find('a', attrs = {'target' : '_blank'})
									ws.cell((c, cNB_COHOTE)).value = 2
								except:
									try:
										the_tr1= the_tr33.find('li', attrs = {'class' : '_108byt5'})
										tt11= the_tr1.find('a', attrs = {'target' : '_blank'})
										ws.cell((c, cNB_COHOTE)).value = 1
									except:
										try:
											the_tr1= the_tr33.findAll('li', attrs = {'class' : '_1rzj9z6'})[0]
											tt11= the_tr1.find('a', attrs = {'target' : '_blank'})
											ws.cell((c, cNB_COHOTE)).value = 2
										except:
											try:
												
												the_tr1= the_tr33.find('li', attrs = {'class' : '_1rzj9z6'})
												tt11= the_tr1.find('a', attrs = {'target' : '_blank'})
												ws.cell((c, cNB_COHOTE)).value = 1
											except:
												aaa=1
								try:
									tt11= the_tr1.find('a', attrs = {'target' : '_blank'})
									div1=tt11['href']  #.attrs['href']
									ws.cell((c, cCOHOTE_URL1)).value = "https://www.airbnb.fr"+str(div1)
									tt12= the_tr1.find('span', attrs = {'class' : '_1kfl0pr'})
									ws.cell((c, cCOHOTE_NAME1)).value = tt12.text
									try:
										tt13= the_tr1.find('img', attrs = {'class' : '_9ofhsl'})
										ws.cell((c, cCOHOTE_IMAGE1)).value = tt13['src']
									except:
										noimage=1
									the_tr2= the_tr33.findAll('li')[1]
									tt21= the_tr2.find('a', attrs = {'target' : '_blank'})
									div2=tt21['href']  #.attrs['href']
									ws.cell((c, cCOHOTE_URL2)).value = "https://www.airbnb.fr"+str(div2)
									tt22= the_tr2.find('span', attrs = {'class' : '_1kfl0pr'})
									ws.cell((c, cCOHOTE_NAME2)).value = tt22.text
									tt23= the_tr2.find('img', attrs = {'class' : '_9ofhsl'})
									ws.cell((c, cCOHOTE_IMAGE2)).value = tt23['src']
								except:
									aaa=1
						#PROPRETE
							try:
								tt= soup.findAll('span', attrs={"class": "_4oybiu"})[0]
								#print(tt.text)
								ws.cell((c, cPROPRETE)).value = tt.text
							except:
								#print('no proprete')
								aaa=1
						#PRECISION
							try:
								tt= soup.findAll('span', attrs={"class": "_4oybiu"})[1]
								#print(tt.text)
								ws.cell((c, cPRECISION)).value = tt.text
							except:
								#print('no Precision')
								aaa=1
						#COMMUNICATION
							try:
								tt= soup.findAll('span', attrs={"class": "_4oybiu"})[2]
								#print(tt.text)
								ws.cell((c, cCOMMUNICATION)).value = tt.text
							except:
								#print('no communication')
								aaa=1
						#EMPLACEMENT
							try:
								tt= soup.findAll('span', attrs={"class": "_4oybiu"})[3]
								#print(tt.text)
								ws.cell((c, cEMPLACEMENT)).value = tt.text
							except:
								#print('no emplacement')
								aaa=1
						#ARRIVEE
							try:
								tt= soup.findAll('span', attrs={"class": "_4oybiu"})[4]
								#print(tt.text)
								ws.cell((c, cARRIVEE)).value = tt.text
							except:
								#print('no arrivee')
								aaa=1
						#QUALITY PRICE
							try:
								tt= soup.findAll('span', attrs={"class": "_4oybiu"})[5]
								#print(tt.text)
								ws.cell((c, cQUALITY_PRICE)).value = tt.text
							except:
								#print('no price quality')
								aaa=1
					#N° ENREGISTREMENT
							try:
								the_tr= soup.find('li', text=re.compile(r'\bNuméro\b'), attrs = {'class' : '_1q2lt74'})
								pp=the_tr.text
								print("1-"+str(pp))
								sp=pp.split(' ')
								ws.cell((c, cREGISTER)).value = sp[-1]
							except:
								try:
									the_tr= soup.find('ul', attrs = {'class' : 'fhhmddr dir dir-ltr'})
									#pp= the_tr.find('li', text=re.compile(r'\bNuméro\b'))
									pp= the_tr.findAll('li')[0]
									if "Numéro" in pp.text:
										txt=pp.span.text
										ws.cell((c, cREGISTER)).value = txt
										print("2-"+str(txt))
								except:
									try:
										the_tr= soup.find('li', text=re.compile(r'\bNuméro\b'), attrs = {'class' : 'f19phm7j dir dir-ltr'})
										pp=the_tr.text
										print("3-"+str(pp))
										sp=pp.split(' ')
										ws.cell((c, cREGISTER)).value = sp[-1]
									except:
										aaa=1
					#TAUX REPONSE
							try:
								the_tr=soup.find('li', text=re.compile(r'\bTaux\b'))
								pp=the_tr.text
								pp=pp.replace(" ","")
								#print(pp)
								sp=pp.split(':')
								#print(sp[-1])
								ws.cell((c, cTAUX_REPONSE)).value = sp[-1]
							except:
								#print('no taux réponse')
								aaa=1
					#DELAI REPONSE
							try:
								the_tr=soup.find('li', text=re.compile(r'\bDélai\b'))
								pp=the_tr.text
								sp=pp.split(':')
								ws.cell((c, cDELAI_REPONSE)).value = sp[-1]
							except:
								#print('no DELAI REPONSE')
								aaa=1
					#DURING SEJOUR
							try:
								the_tr=soup.findAll('div', attrs={"class": "ciubx2o dir dir-ltr"})[-1]
							except:
								try:
									the_tr=soup.findAll('div', attrs={"class": "_1byskwn"})[-1]
								except:
									aaa=1
								aaa=1

							try:
								tt= the_tr.find('span', text=re.compile(r'\bArrivée\b'))
								ws.cell((c, cCHECK_IN)).value = tt.text
								#print(tt.text)
							except:
								#print('no ARRIVE')
								aaa=1
							try:
								tt= the_tr.find('span', text=re.compile(r'\bDépart\b'))
								ws.cell((c, cCHECK_OUT)).value = tt.text
							except:
								#print('no DEPART')
								aaa=1
							try:
								tt= the_tr.find('span', text=re.compile(r'\bNon fumeur\b'))
								ws.cell((c, cFUMEUR)).value = tt.text
							except:
								#print('no FUMEUR')
								aaa=1
							try:
								tt= the_tr.find('span', text=re.compile(r'\bNe convient pas aux\b'))
								ws.cell((c, cENFANT)).value = tt.text
							except:
								#print('no CHILD')
								aaa=1
							try:
								tt= the_tr.find('span', text=re.compile(r"\bArrivée autonome\b"))
								ws.cell((c, cSERRURE)).value = tt.text
							except:
								#print('no AUTOMATIC')
								aaa=1
							try:
								tt= the_tr.find('span', text=re.compile(r"\bPas d'animaux\b"))
								ws.cell((c, cANIMAUX)).value = tt.text
							except:
								try:
									tt= the_tr.find('span', text=re.compile(r"\bAnimaux de compagnie\b"))
									ws.cell((c, cANIMAUX)).value = tt.text
								except:
									#print('no ANIMAL')
									aaa=1
							try:
								tt= the_tr.find('span', text=re.compile(r"\bCaution\b"))
								ws.cell((c, cCAUTION)).value = tt.text
							except:
								#print('no Caution')
								aaa=1
							try:
								tt= the_tr.find('span', text=re.compile(r"\bDétecteur de fumée\b"))
								ws.cell((c, cFUMEE)).value = tt.text
							except:
								#print('no detecteur fumee')
								aaa=1
							try:
								tt= the_tr.find('span', text=re.compile(r"\bDétecteur de monoxyde de carbone\b"))
								ws.cell((c, cMONOXYDE)).value = tt.text
							except:
								#print('no detecteur monoxyde')
								aaa=1
							try:
								tt= the_tr.find('span', text=re.compile(r"\bPas de fête ni de soirée\b"))
								ws.cell((c, cFETE)).value = tt.text
							except:
								#print('no detecteur monoxyde')
								aaa=1
							try:
								tt= the_tr.find('span', text=re.compile(r"\bmatière de distanciation sociale\b"))
								ws.cell((c, cDISTANCIATION_SOCIAL)).value = 'Y'
							except:
								ws.cell((c, cDISTANCIATION_SOCIAL)).value = 'N'
								#print('no distanciation sociale')
					#LANGUE
							try:
								the_tr= soup.find('li', text=re.compile(r'\bLangues\b'))
								#print(the_tr)
								pp=the_tr.text
								#print(pp)
								sp=pp.split(':')
								ws.cell((c, cLANGUE)).value = sp[-1]
							except:
								#print('no LANGUAGE')
								aaa=1
					#IMAGE
							try:
								the_tr= soup.findAll('img', attrs={"class": "_6tbg2q"})[0]
								#print(the_tr)
								tt=the_tr['src']
								ws.cell((c, cIMAGE_1)).value = tt
							except:
								try:
									the_tr= soup.find('img', attrs={"class": "_9ofhsl"})
									#print(the_tr)
									tt=the_tr['data-original-uri']
									ws.cell((c, cIMAGE_1)).value = tt
								except:
									#print('no IMAGE 0')
									aaa=1
							try:
								the_tr= soup.findAll('img', attrs={"class": "_6tbg2q"})[1]
								#print(the_tr)
								tt=the_tr['src']
								ws.cell((c, cIMAGE_2)).value = tt
							except:
								try:
									the_tr= soup.findAll('img', attrs={"class": "_9ofhsl"})[1]
									#print(the_tr)
									tt=the_tr['data-original-uri']
									ws.cell((c, cIMAGE_2)).value = tt
								except:
									#print('no IMAGE 1')
									aaa=1
							try:
								the_tr= soup.findAll('img', attrs={"class": "_6tbg2q"})[2]
								#print(the_tr)
								tt=the_tr['src']
								ws.cell((c, cIMAGE_3)).value = tt
							except:
								try:
									the_tr= soup.findAll('img', attrs={"class": "_9ofhsl"})[2]
									#print(the_tr)
									tt=the_tr['data-original-uri']
									ws.cell((c, cIMAGE_3)).value = tt
								except:
									#print('no IMAGE 2')
									aaa=1
							try:
								the_tr= soup.findAll('img', attrs={"class": "_6tbg2q"})[3]
								#print(the_tr)
								tt=the_tr['src']
								ws.cell((c, cIMAGE_4)).value = tt
							except:
								try:
									the_tr= soup.findAll('img', attrs={"class": "_9ofhsl"})[3]
									#print(the_tr)
									tt=the_tr['data-original-uri']
									ws.cell((c, cIMAGE_4)).value = tt
								except:
									#print('no IMAGE 3')
									aaa=1
							try:
								the_tr= soup.findAll('img', attrs={"class": "_6tbg2q"})[4]
								#print(the_tr)
								tt=the_tr['src']
								#print(tt)
								ws.cell((c, cIMAGE_5)).value = tt
							except:
								try:
									the_tr= soup.findAll('img', attrs={"class": "_9ofhsl"})[4]
									#print(the_tr)
									tt=the_tr['data-original-uri']
									ws.cell((c, cIMAGE_5)).value = tt
								except:
									#print('no IMAGE 4')
									aaa=1
					#IMAGE_HOTE
							try:
								the_tr= soup.find('div', attrs={"class": "_5kripx"})
								t= the_tr.find('img', attrs={"class": "_9ofhsl"})
								tt=t['src']
								ws.cell((c, cIMAGE_PROFIL)).value = tt
							except:
								#print('no IMAGE_HOTE')
								aaa=1
					#ENTREPRISE
							try:
								the_tr= soup.findAll('ol', attrs={"class": "lgx66tx dir dir-ltr"})[-1]
								t= the_tr.find('button')
								if "Professionnel" in t.text:
									ws.cell((c, cENTREPRISE)).value = 'YES'
									#print('ENTREPRISE YES')
								else:
									ws.cell((c, cENTREPRISE)).value = 'NO'
							except:
								#print('no IMAGE_HOTE')
								ws.cell((c, cENTREPRISE)).value = 'NO'
								aaa=1
					#SOCIALWASHING
							try:
								t0=[]
								t1=[]
								t2=[]
								try:
									t0= soup.findAll('span', attrs = {'class' : '_pog3hg'})[0]
								except:
									pass
								try:
									t1= soup.findAll('span', attrs = {'class' : '_pog3hg'})[1]
								except:
									pass
								try:
									t2= soup.findAll('span', attrs = {'class' : '_pog3hg'})[2]
								except:
									pass
								try:
									t3= soup.findAll('span', attrs = {'class' : '_pog3hg'})[3]
								except:
									pass
								if "Soutien" in t0.text:
									ws.cell((c, cSOCIALWASHING)).value = 'YES'
								elif "Soutien" in t1.text:
									ws.cell((c, cSOCIALWASHING)).value = 'YES'
								elif "Soutien" in t2.text:
									ws.cell((c, cSOCIALWASHING)).value = 'YES'
								elif "Soutien" in t3.text:
									ws.cell((c, cSOCIALWASHING)).value = 'YES'
								else:
									ws.cell((c, cSOCIALWASHING)).value = 'NO'
							except:
								ws.cell((c, cSOCIALWASHING)).value = 'NO'
								aaa=1
					#ADRESSE
							try:
								t= soup.find('div', attrs={"data-plugin-in-point-id": "LOCATION_DEFAULT"})								
								tt= t.find('div', attrs={"class": "_152qbzi"})
								ws.cell((c, cADRESS)).value = tt
								#print(tt)
							except:
								aaa=1
							ws.cell((c, cACTIVE)).value = 'YES'
			#------------------------
					except:
						pass

				except:
					pass
				timer=1000
			else:
				time.sleep(1)
				timer=timer+1
		if timer==61:
			driver.quit()


	c=c+1
print ('_______    ___    ___     ___')
print ('|      |   |  |   |  \    |  |')
print ('|  |__     |  |   |   \   |  |')
print ('|     |    |  |   |    \  |  |')
print ('|  |       |  |   |  |\ \ |  |')
print ('|  |       |  |   |  | \ \|  |')
print ('|__|       |__|   |__|  \____|')

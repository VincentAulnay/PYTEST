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

def GSwrite(c):
	print('=======test écriture title')
	ws.cell((c, 2)).value = 'test'
	print('=======print ok')
	
#threading.Thread(target=GSwrite, args=(c, clevel, valeur,)).start()
	
fm=2
fff=0

nrow=100000
h=ws.cell((c, cANNONCE)).value
threading.Thread(target=scrap, args=(h,)).start()
scrap_ok=1
time.sleep(8)
while c<=nrow:
	#scrap_ok=0
	print (c)
	if (c/1000).is_integer():
		driver = webdriver.Chrome('/usr/lib/chromium-browser/chromedriver',chrome_options=chrome_options)
		driver.set_window_size(2000, 1000)
		driver.get('https://www.google.com/')
		driver.get('chrome://settings/')
		driver.execute_script('chrome.settingsPrivate.setDefaultZoom(0.5);')
		driver.implicitly_wait(10)
		wait = WebDriverWait(driver, 5)
		wait2 = WebDriverWait(driver, 5)
		wait3 = WebDriverWait(driver, 5)
		wait = WebDriverWait(driver, 2)
		time.sleep(2)
		driver.get(h)
		time.sleep(2)
		scrap_ok=1
	#h=ws.cell((c, cANNONCE)).value
	#driver.get(h)
	#time.sleep(5)
	numero=None
	#print (h)
	#do=sheet_read.cell(i,0).value
	if numero is None:
		#threading.Thread(target=scrap, args=(h,)).start()
		timer=1
		while timer<=60:
			if scrap_ok==1:
				try:
					f_ele=0
					if fm==2:
						try:
							button_fermer = wait.until(EC.presence_of_element_located((By.XPATH, "//button[@aria-label='Fermer']")))
							button_fermer.click()
						except:
							pass
						while f_ele<=3:
							try:
								#ele=driver.find_element_by_xpath("//div[@class='_1cvivhm']")
								#ele=driver.find_element_by_xpath("//div[@class='_cg8a3u']")
								driver.execute_script("window.scrollBy(0,3000);")
								ele=driver.find_element_by_xpath("//div[@class='s9fngse dir dir-ltr']")
								driver.execute_script("arguments[0].scrollIntoView(true);", ele)
								#driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
								#time.sleep(1)
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
					h=ws.cell((c+1, cANNONCE)).value
					FTitle = soup.find('div', attrs={"data-plugin-in-point-id": "TITLE_DEFAULT"})
					Flogement = soup.find('div', attrs={"data-plugin-in-point-id": "OVERVIEW_DEFAULT"})
					FProfile = soup.find('div', attrs={"data-plugin-in-point-id": "HOST_PROFILE_DEFAULT"})
					FPolicies = soup.find('div', attrs={"data-plugin-in-point-id": "POLICIES_DEFAULT"})
					FHero = soup.find('div', attrs={"data-plugin-in-point-id": "HERO_DEFAULT"})
					h=ws.cell((c+1, cANNONCE)).value
					threading.Thread(target=scrap, args=(h,)).start()
					#print('start bs4')
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
								ws.cell((c, clat)).value = long_lat[0]
								#print('try threading')
								#threading.Thread(target=GSwrite, args=(c,clat,long_lat[0],)).start()
								#run_write=GSWrite(c,clat,long_lat[0])
								#print('done')
								ws.cell((c, clon)).value = long_lat[1]
								#threading.Thread(target=GSwrite, args=(c,clon,long_lat[1],)).start()
								#run_write=GSWrite(c,clon,long_lat[1])
							except:
								ee=1
						#TITLE
							try:
								div1=FTitle.find('h1', attrs={"class": "_fecoyn4"})
								ws.cell((c, cTITLE)).value = div1.text
								#threading.Thread(target=scrap, args=(h,)).start()
								#threading.Thread(target=GSwrite, args=(c,)).start()
								#run_write=GSWrite(c,cTITLE,div1.text)
								#print(div1.text)
							except:
								#print('NO TITLE')
								aaa=1
						#URL HOTE
							try:
								div=FProfile.find('div', attrs={"class": "c6y5den dir dir-ltr"})
								div2=div.find('a')
								div1=div2['href']  #.attrs['href']
								ws.cell((c, cHOTE)).value = "https://www.airbnb.fr"+str(div1)
								#print("URLHOT1"+str(div1))
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
									tp_c=FTitle.find('span', attrs={"class": "_s65ijh7"}).text
									#print(tp_c)
									#print('type1')

								except:
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
							except:
								#print('NOCOMMENT')
								aaa=1
						#VOYAGEUR
							try:
								#the_tr= Flogement.find('div', attrs = {'class' : '_tqmy57'})
								tt=Flogement.find_all('li')[0]
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
								#the_tr= Flogement.find('div', attrs = {'class' : '_tqmy57'})
								tt=Flogement.find_all('li')[2]
								tt1=tt.find_all('span')[2]
								tt2=tt1.text
								p_tp=tt2.split(" ")
								ws.cell((c, cLITS)).value = p_tp[0]
								#print("L="+str(p_tp[0]))
							except:
								#print('NO LIT')
								aaa=1
						#SdB
							try:
								#the_tr= soup.find('div', attrs = {'class' : '_tqmy57'})
								tt=Flogement.find_all('li')[3]
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
								#the_tr= soup.find('div', attrs = {'class' : '_tqmy57'})
								tt=Flogement.find_all('li')[1]
								tt1=tt.find_all('span')[2]
								tt2=tt1.text
								p_tp=tt2.split(" ")
								ws.cell((c, cCHAMBRE)).value = p_tp[0]
								#print("C="+str(p_tp[0]))
							except:
								#print('NO CHAMBRE')
								aaa=1
						#VILLE
							try:
								tp_c=FTitle.find('span', attrs={"class": "_9xiloll"}).text
								ws.cell((c, cVILLE)).value = tp_c
								#print(tp_c)
							except:
								print('NO VILLE')
								aaa=1

							#print(tt)
						#NAME_HOTE
							try:
								tp_c=FProfile.find_all('div', attrs={"class": "hnwb2pb dir dir-ltr"})[0].text
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
										#tp_c=FProfile.find_all('h2', attrs={"class": "hnwb2pb dir dir-ltr"})[1].text
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
								the_tr= Flogement.find('div', attrs={"class": "_cv5qq4"})
								ttt=the_tr.h2.text
								pp=ttt.split('⸱')
								#print(pp)
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
								tp_c=FProfile.find('div', attrs={"class": "s9fngse dir dir-ltr"}).text
								pp=tp_c.split("depuis")
								ws.cell((c, cANCIENNETE)).value = pp[1]
								#print(tp_c)
							except:
								try:
									tp_c=FProfile.find('div', attrs={"class": "_f47qa6"})
									tt=tp_c.find('div', attrs={"class": "_svr7sj"})
									tt1=tt.div.get_text()
									ws.cell((c, cANCIENNETE)).value = tt1
								except:
									aaa=1
						#SUPER HOTE
							try:
								#the_tr= FProfile.find('span', text=re.compile(r'\bSuperhost\b'),attrs = {'aria-hidden' : 'false'})
								#tp_c=soup.find('span', attrs={"class": "_63km3vu"}, text=re.compile(r'\bSuper\b'))
								tp_c=FTitle.find('span', attrs={"class": "_1mhorg9"})
								if tp_c is not None:
									ws.cell((c, cSUPERHOTE)).value = 'X'
									#print (tp_c)
							except:
								aaa=1
						#COMMENT PROFIL

							try:
								tp_c=FProfile.find('ul', attrs={"class": "tq6hspd h1aqtv1m dir dir-ltr"})
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

					#N° ENREGISTREMENT
							try:
								the_tr= FProfile.find('ul', attrs = {'class' : 'fhhmddr dir dir-ltr'})
								#pp= the_tr.find('li', text=re.compile(r'\bNuméro\b'))
								pp= the_tr.findAll('li')[0]
								if "Numéro" in pp.text:
									txt=pp.span.text
									ws.cell((c, cREGISTER)).value = txt
									print("2-"+str(txt))
							except:
								aaa=1

					#DURING SEJOUR
							try:
								tt= FPolicies.find('span', text=re.compile(r'\bArrivée\b'))
								ws.cell((c, cCHECK_IN)).value = tt.text
								#print(tt.text)
							except:
								#print('no ARRIVE')
								aaa=1
							try:
								tt= FPolicies.find('span', text=re.compile(r'\bDépart\b'))
								ws.cell((c, cCHECK_OUT)).value = tt.text
							except:
								#print('no DEPART')
								aaa=1
							try:
								tt= FPolicies.find('span', text=re.compile(r'\bNon fumeur\b'))
								ws.cell((c, cFUMEUR)).value = tt.text
							except:
								#print('no FUMEUR')
								aaa=1
							try:
								tt= FPolicies.find('span', text=re.compile(r'\bNe convient pas aux\b'))
								ws.cell((c, cENFANT)).value = tt.text
							except:
								#print('no CHILD')
								aaa=1
							try:
								tt= FPolicies.find('span', text=re.compile(r"\bArrivée autonome\b"))
								ws.cell((c, cSERRURE)).value = tt.text
							except:
								#print('no AUTOMATIC')
								aaa=1
							try:
								tt= FPolicies.find('span', text=re.compile(r"\bPas d'animaux\b"))
								ws.cell((c, cANIMAUX)).value = tt.text
							except:
								try:
									tt= FPolicies.find('span', text=re.compile(r"\bAnimaux de compagnie\b"))
									ws.cell((c, cANIMAUX)).value = tt.text
								except:
									#print('no ANIMAL')
									aaa=1
							try:
								tt= FPolicies.find('span', text=re.compile(r"\bDétecteur de fumée\b"))
								ws.cell((c, cFUMEE)).value = tt.text
							except:
								#print('no detecteur fumee')
								aaa=1
							try:
								tt= FPolicies.find('span', text=re.compile(r"\bDétecteur de monoxyde de carbone\b"))
								ws.cell((c, cMONOXYDE)).value = tt.text
							except:
								#print('no detecteur monoxyde')
								aaa=1
							try:
								tt= FPolicies.find('span', text=re.compile(r"\bPas de fête ni de soirée\b"))
								ws.cell((c, cFETE)).value = tt.text
							except:
								#print('no detecteur monoxyde')
								aaa=1
					#IMAGE
							try:
								the_tr= FHero.findAll('img', attrs={"class": "_6tbg2q"})[0]
								#print(the_tr)
								tt=the_tr['src']
								ws.cell((c, cIMAGE_1)).value = tt
							except:
								aaa=1
							try:
								the_tr= FHero.findAll('img', attrs={"class": "_6tbg2q"})[1]
								#print(the_tr)
								tt=the_tr['src']
								ws.cell((c, cIMAGE_2)).value = tt
							except:
								aaa=1
							try:
								the_tr= FHero.findAll('img', attrs={"class": "_6tbg2q"})[2]
								#print(the_tr)
								tt=the_tr['src']
								ws.cell((c, cIMAGE_3)).value = tt
							except:
								aaa=1
							try:
								the_tr= FHero.findAll('img', attrs={"class": "_6tbg2q"})[3]
								#print(the_tr)
								tt=the_tr['src']
								ws.cell((c, cIMAGE_4)).value = tt
							except:
								aaa=1
							try:
								the_tr= FHero.findAll('img', attrs={"class": "_6tbg2q"})[4]
								#print(the_tr)
								tt=the_tr['src']
								#print(tt)
								ws.cell((c, cIMAGE_5)).value = tt
							except:
								aaa=1
					#IMAGE_HOTE
							try:
								#the_tr= soup.find('div', attrs={"class": "_5kripx"})
								t= Flogement.find('img', attrs={"class": "_9ofhsl"})
								tt=t['src']
								ws.cell((c, cIMAGE_PROFIL)).value = tt
							except:
								#print('no IMAGE_HOTE')
								aaa=1
					#ENTREPRISE
							try:
								the_tr= FProfile.findAll('ol', attrs={"class": "lgx66tx dir dir-ltr"})[-1]
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
			driver = webdriver.Chrome('/usr/lib/chromium-browser/chromedriver',chrome_options=chrome_options)
			driver.set_window_size(2000, 1000)
			driver.get('https://www.google.com/')
			driver.get('chrome://settings/')
			driver.execute_script('chrome.settingsPrivate.setDefaultZoom(0.5);')
			driver.implicitly_wait(10)
			wait = WebDriverWait(driver, 5)
			wait2 = WebDriverWait(driver, 5)
			wait3 = WebDriverWait(driver, 5)
			wait = WebDriverWait(driver, 2)
			time.sleep(2)
			driver.get(h)
			time.sleep(2)
			scrap_ok=1


	c=c+1
print ('_______    ___    ___     ___')
print ('|      |   |  |   |  \    |  |')
print ('|  |__     |  |   |   \   |  |')
print ('|     |    |  |   |    \  |  |')
print ('|  |       |  |   |  |\ \ |  |')
print ('|  |       |  |   |  | \ \|  |')
print ('|__|       |__|   |__|  \____|')

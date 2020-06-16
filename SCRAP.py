from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import xlwt
import xlrd
import time
from xlutils.copy import copy
import re
import json
from urllib.request import urlopen
from openpyxl import load_workbook
from tkinter import filedialog
from tkinter import *
import os
from bs4 import BeautifulSoup
import threading
import sys


chrome_options = webdriver.ChromeOptions()
#prefs = {"profile.managed_default_content_settings.images": 2}
#chrome_options.add_experimental_option("prefs", prefs)

#driver.set_window_position(0, 0)
print ('▀▄▀▄▀▄ STOP AIRBNB ▄▀▄▀▄▀')

#--------SELECTION DU FICHIER AVEC LES NOUVELLES ANNONCES--------

wbx = load_workbook(path_RESULT.filename)
ws = wbx.active
#INITIALISATION EXCEL


#Récupération des URL depuis le xls
#list_URL=sheet_read.col_values(1)
list_URL=[]
for col in ws['B']:
    list_URL.append(col.value)

del list_URL[0]

#Récupérer les XPATH

YN_title='YES'
YN_profil='YES'
YN_name='YES'
YN_price='NO'
YN_comment='YES'
YN_voyageur='YES'
YN_chamber='YES'
YN_bed='YES'
YN_gps='YES'
YN_type='YES'
YN_old1='YES'
YN_old2='YES'
YN_ville='YES'
YN_super='YES'
YN_SdB='YES'

#OUVERTURE DES PAGES CHROME
#driver = webdriver.Chrome(chrome_options=chrome_options)
driver = webdriver.Chrome('/usr/lib/chromium-browser/chromedriver')
#driver = webdriver.Chrome()
driver.set_window_size(800, 800)
time.sleep(2)
#c = ligne 2 du xls resultant
c=2
wait2 = WebDriverWait(driver, 2)
wait3 = WebDriverWait(driver, 3)	

i=1

def extract(XP, nb, type, c, YP):
	wait = WebDriverWait(driver, 2)
	if YP=='YES':
		try:
			tp = wait.until(EC.presence_of_element_located((By.XPATH, XP))).text
			#sheet.write(c, nb, tp)
			ws.cell(row=c+1, column=nb+1).value = tp
			print (tp)
		except:
			print(type)
			pass
end=0
while end==0:
	try:
		for h in list_URL:
			print (c)
			print (h)
			hh=ws.cell(row=c, column=2).value
			#do=sheet_read.cell(i,0).value
			do=ws.cell(row=c, column=1).value
			if do==None:
				driver.get(hh)
				time.sleep(4)
				f_ele=0
				while f_ele<=3:
					try:
						ele=driver.find_element_by_xpath("//div[@class='_384m8u']")
						driver.execute_script("arguments[0].scrollIntoView(true);", ele)
						driver.execute_script("window.scrollBy(0,-100);")
						#driver.execute_script("window.scrollBy(0,500);")
						f_ele=6
						time.sleep(3)
					except:
						f_ele=f_ele+1
						time.sleep(1)
			#PROFILE
				html = driver.page_source
				time.sleep(1)
				soup = BeautifulSoup(html, 'html.parser')
				time.sleep(2)
				try:
				#TITLE
					try:
						div1=soup.find('div', attrs={"class": "_5z4v7g"})
						ws.cell(row=c, column=1).value = div1.h1.text
					except:
						print('NO TITLE')
				#URL HOTE
					try:
						div=soup.findAll('a', attrs={"class": "_105023be"})[-1]
						div1=div['href']  #.attrs['href']
						ws.cell(row=c, column=5).value = "https://www.airbnb.fr"+str(div1)
					except:
						print('NO PROFILE')
				#COMMENTAIRE
					COMMENT='NO COMMENT'
					#run_price=extract("//span[@class='_wfad8t']",6,COMMENT,c,YN_comment)
					try:
						p_c=[]
						tp_c=soup.findAll('span', attrs={"class": "_bq6krt"})[1].text
						p_c=tp_c.replace("(","")
						cc=p_c.replace(")","")
						try:
							pp=cc.split(' ')
							cc=pp[1]
						except:
							pass
						ws.cell(row=c, column=7).value = cc
						print ("COMMENT ===")
						print(cc)
						#p_c=tp_c.split("(")
						#print('ici1')
						#table_c = p_c[1].replace(")"," ")
						#print (table_c)
						#ws.cell(row=c+1, column=6+1).value = table_c
					except:
						print('NOCOMMENT')
				#VOYAGEUR
					if YN_voyageur=='YES':
						try:
							the_tr= soup.find('div', attrs = {'class' : '_tqmy57'})
							tt=the_tr.find_all('div')[1]
							tt1=tt.find_all('span')[0]
							tt2=tt1.text
							p_tp=tt2.split(" ")
							ws.cell(row=c, column=9).value = p_tp[0]
						except:
							print('NO VOYAGER')

				#LITS
					if YN_bed=='YES':
						try:
							the_tr= soup.find('div', attrs = {'class' : '_tqmy57'})
							tt=the_tr.find_all('div')[1]
							tt1=tt.find_all('span')[4]
							tt2=tt1.text
							p_tp=tt2.split(" ")
							ws.cell(row=c, column=12).value = p_tp[0]
						except:
							print('NO LIT')
				#SdB
					if YN_SdB=='YES':
						try:
							the_tr= soup.find('div', attrs = {'class' : '_tqmy57'})
							tt=the_tr.find_all('div')[1]
							tt1=tt.find_all('span')[6]
							tt2=tt1.text
							p_tp=tt2.split(" ")
							ws.cell(row=c, column=11).value = p_tp[0]
						except:
							print('NO SdB')
				#CHAMBRE
					if YN_chamber=='YES':
						try:
							the_tr= soup.find('div', attrs = {'class' : '_tqmy57'})
							tt=the_tr.find_all('div')[1]
							tt1=tt.find_all('span')[2]
							tt2=tt1.text
							p_tp=tt2.split(" ")
							ws.cell(row=c, column=10).value = p_tp[0]
						except:
							print('NO CHAMBRE')
				#VILLE
					try:
						tp_c=soup.find('a', attrs={"class": "_5twioja"}).text
						ws.cell(row=c, column=13).value = tp_c
					except:
						print('NO VILLE')

				#GPS
					try:
						tp_c=soup.find('div', attrs={"class": "gm-style"})
						tt=tp_c.find('div', attrs={"style": "margin-left: 5px; margin-right: 5px; z-index: 1000000; position: absolute; left: 0px; bottom: 0px;"})
						tt1=tt.find('a', attrs={"target": "_blank"})
						tt2=tt1['href']
						#----------Create translation table----------
						table = str.maketrans('=&', '++')
						result_gps = tt2.translate(table)
						split_gps=result_gps.split("+")
						#https://www.google.com/maps?ll+46.23657,3.92341+z=14&t=m&hl=fr&gl=FR&mapclient=apiv3
						#https://maps.google.com/maps?ll=49.19054,-2.11715&z=14&t=m&hl=fr&gl=FR&mapclient=apiv3
						coor=split_gps[1]
						long_lat=coor.split(',')
						#--------------Write results--------------
						print(long_lat[0])
						print(long_lat[1])
						#sheet.write(c, 12, long_lat[0])
						#sheet.write(c, 13, long_lat[1])
						ws.cell(row=c, column=14).value = long_lat[0]
						ws.cell(row=c, column=15).value = long_lat[1]
					except:
						print('NO GPS')
				#NAME_HOTE
					try:
						tp_c=soup.find('div', attrs={"class": "_f47qa6"})
						tt=tp_c.find('div', attrs={"class": "_svr7sj"})
						tt1=tt.h2.get_text()
						pp=tt1.split('par ')
						ws.cell(row=c, column=3).value = pp[1]
					except:
						print ('NO_NAME')
				#TYPE_HOME
					try:
						the_tr= soup.find('div', attrs={"class": "_1b3ij9t"}).text
						pp=the_tr.split('.')
						ws.cell(row=c, column=8).value = pp[0]
					except:
						print('NOTYPE')
					
				#ANCIENNETE
					try:
						tp_c=soup.find('div', attrs={"class": "_f47qa6"})
						tt=tp_c.find('div', attrs={"class": "_svr7sj"})
						tt1=tt.div.get_text()
						ws.cell(row=c, column=4).value = tt1
					except:
						print ('NOOLD')

				#SUPER HOTE
					try:
						#the_tr= soup.find('span', text=re.compile(r'\bSuperhost\b'),attrs = {'aria-hidden' : 'false'})
						tp_c=soup.find('div', attrs={"class": "_1ft6jxp"}).text
						ws.cell(row=c, column=16).value = 'X'
					except:
						print('no superhote')
				#AUTONOME
					try:
						the_tr= soup.find('div', text=re.compile(r'\bArrivée autonome\b'),attrs = {'class' : '_1qsawv5'})
						div2=the_tr.findNextSibling('div')
						print(the_tr.text)
						ws.cell(row=c, column=17).value = div2.text
						print(div2.text)
					except:
						print('no auto')
				#CHILDREN
					try:
						the_tr= soup.find('span', text=re.compile(r'\bNe convient pas aux enfants ni aux bébés\b'))
						ws.cell(row=c, column=18).value = the_tr.text
					except:
						print('no child')
				#ANNULATION
					try:
						the_tr= soup.find('div', text=re.compile(r'\bAnnulation gratuite\b'),attrs = {'class' : '_1qsawv5'})
						div2=the_tr.findNextSibling('div')
						ws.cell(row=c, column=19).value = div2.text
					except:
						print('no child')

					if (c/10).is_integer():
						wbx.save(path_RESULT.filename)
					if (c/1000).is_integer():
						wbx.save(DIR2+NAMEFile+str(c)+".xlsx")
#------------------------
				except:
					try:
						#page sans annonce, je dois trouver quelque soit le type
						wait3.until(EC.presence_of_element_located((By.XPATH, "//h3[@class='_jmmm34f']")))
					except:
						# rien tourvé précédent, donc c'est que je suis sur mauvais design, clos et reopen chrome
						driver.quit()
						f_xpathdate=0
						fff=0
						fm=2
						driver = webdriver.Chrome('/usr/lib/chromium-browser/chromedriver')
						#driver = webdriver.Chrome()
						driver.set_window_size(800, 800)
						wait3 = WebDriverWait(driver, 3)
						while f_xpathdate==0:
							h=ws.cell(row=fm, column=2).value
							print(h)
							if fff==10:
								f_mounth=1
								f_xpathdate=1
								end=0
							fff=fff+1
							try:
								driver.get(h)
								time.sleep(4)
								x_date = wait3.until(EC.presence_of_element_located((By.XPATH, "//div[@class='_13m7kz7i']"))).text
								print("x date trouve")
								f_xpathdate=1
							except:
								if fff!=10:
									driver.quit()
									#driver = webdriver.Chrome('/usr/lib/chromium-browser/chromedriver')
									driver = webdriver.Chrome()
									driver.set_window_size(800, 800)
									wait3 = WebDriverWait(driver, 3)

			c=c+1
					

	except:
		print("END")
		# EXCEPT si Chrome se ferme tout seul, ici il va le réouvrir et relancer la boucle d'extraction
		#driver = webdriver.Chrome()
		#driver.set_window_size(800, 1500)
		wbx.save(path_RESULT.filename)
	wbx.save(path_RESULT.filename)
	print ('_______    ___    ___     ___')
	print ('|      |   |  |   |  \    |  |')
	print ('|  |__     |  |   |   \   |  |')
	print ('|     |    |  |   |    \  |  |')
	print ('|  |       |  |   |  |\ \ |  |')
	print ('|  |       |  |   |  | \ \|  |')
	print ('|__|       |__|   |__|  \____|')
	print ('Votre fichier RESULT est à présent terminé, renommer le et concerver le.')
	print ('Vous pouvez maintenant:')
	print ('  1- exécuter le .exe N°4 qui calcule le nombre de logement détenus par hôte')
	print ('  2- exétuter le .exe N°5 qui lui calcule les nuitées de chaque annonces, à exécuter hebdomadairement')
	print ('  3- N°6 qui est à utiliser si votre fichier RESULT contient plus de 2000 annonces, il va découper le fichier en lot de moins de 2000 annonces si vous voulez l importer sur GoogleMap')
	print ('  4- N°7 vous permet de comparer le fichier RESULT avec un autre fichier de liste d annonce que vous possédez déjà, il va alors ajouter dans votre autre fichier les annonces manquante')
	end=1	

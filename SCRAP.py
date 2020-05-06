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
driver = webdriver.Chrome(/usr/lib/chromium-browser/chromedriver')
driver.set_window_size(800, 800)

#c = ligne 2 du xls resultant
c=1
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
			#do=sheet_read.cell(i,0).value
			do=ws.cell(row=i+1, column=1).value
			if do==None:
				driver.get(h)
				time.sleep(2)
				if c==1:
					time.sleep(5)
			#PROFILE
				try:
					#je test si je suis sur une annonce au bon design
					wait3.until(EC.presence_of_element_located((By.XPATH, "//span[@class='_18hrqvin']")))
					#OK j'extrait les détails
#------------------------					
				#URL HOTE
					try:
						element = driver.find_element_by_xpath("//div[@class='_1ij6gln6']/a")
						hote=element.get_attribute('href')
						#sheet.write(c, 4, hote)
						ws.cell(row=c+1, column=4+1).value = hote
						print ("URL HOTE ===")
						print (hote)
					except:
						print('NO PROFILE')
				#PRICE
					#PRICE='NO PRICE'
					#run_price=extract(X_price,5,PRICE,c,YN_price)
				#COMMENTAIRE
					COMMENT='NO COMMENT'
					#run_price=extract("//span[@class='_wfad8t']",6,COMMENT,c,YN_comment)
					try:
						p_c=[]
						tp_c = wait3.until(EC.presence_of_element_located((By.XPATH, "//span[@class='_1plk0jz1']"))).text
						p_c=tp_c.replace("(","")
						cc=p_c.replace(")","")
						Scomment=cc.replace(" ","")
						Lcomment=Scomment.split("c")
						Icomment=int(Lcomment[0])
						ws.cell(row=c+1, column=6+1).value = Icomment
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
							#//*[@id="room"]/div[2]/div/div[2]/div[1]/div/div[3]/div/div/div/div[1]/div[2]/div[2]/div[1]/div
							#//*[@id="room"]/div[2]/div/div[2]/div[1]/div/div[3]/div/div/div/div[1]/div/div/div[1]/div
							#//*[@id="room"]/div[2]/div/div[2]/div[1]/div/div[3]/div/div/div/div[1]/div/div/div[2]/div
							#//*[@id="room"]/div[2]/div/div[2]/div[1]/div/div[3]/div/div/div/div[1]/div[2]/div[2]/div[2]/div
							tp = wait3.until(EC.presence_of_element_located((By.XPATH, "//*[@id='room']/div[2]/div/div[2]/div[1]/div/div[3]/div/div/div[1]/div/div/div[1]/div"))).text
							p_tp=tp.split(" ")
							ws.cell(row=c+1, column=8+1).value = p_tp[0]
							print ("VOYAGEUR ===")
							print (p_tp[0])
						except:
							try:
								tp = wait3.until(EC.presence_of_element_located((By.XPATH, "//*[@id='room']/div[2]/div/div[2]/div[1]/div/div[3]/div/div/div/div[1]/div/div/div[1]/div"))).text
								p_tp=tp.split(" ")
								ws.cell(row=c+1, column=8+1).value = p_tp[0]
								print ("VOYAGEUR ===")
								print (p_tp[0])
							except:
								print('NO VOYAGER')

				#LITS
					if YN_bed=='YES':
						try:
							tp = wait3.until(EC.presence_of_element_located((By.XPATH, "//*[@id='room']/div[2]/div/div[2]/div[1]/div/div[3]/div/div/div[1]/div/div/div[3]/div"))).text
							p_tp=tp.split(" ")
							ws.cell(row=c+1, column=11+1).value = p_tp[0]
							print ("LIT ===")
							print (p_tp[0])
						except:
							try:
								tp = wait3.until(EC.presence_of_element_located((By.XPATH, "//*[@id='room']/div[2]/div/div[2]/div[1]/div/div[3]/div/div/div/div[1]/div/div/div[3]/div"))).text
								p_tp=tp.split(" ")
								ws.cell(row=c+1, column=11+1).value = p_tp[0]
								print ("LIT ===")
								print (p_tp[0])
							except:
								print('NO LIT')
				#SdB
					if YN_SdB=='YES':
						try:
							tp = wait3.until(EC.presence_of_element_located((By.XPATH, "//*[@id='room']/div[2]/div/div[2]/div[1]/div/div[3]/div/div/div[1]/div/div/div[4]/div"))).text
							p_tp=tp.split(" ")
							ws.cell(row=c+1, column=10+1).value = p_tp[0]
							print ("SdB ===")
							print (p_tp[0])
						except:
							try:
								tp = wait3.until(EC.presence_of_element_located((By.XPATH, "//*[@id='room']/div[2]/div/div[2]/div[1]/div/div[3]/div/div/div/div[1]/div/div/div[4]/div"))).text
								p_tp=tp.split(" ")
								ws.cell(row=c+1, column=10+1).value = p_tp[0]
								print ("SdB ===")
								print (p_tp[0])
							except:
								print('NO SdB')
				#CHAMBRE
					if YN_chamber=='YES':
						try:
							tp = wait3.until(EC.presence_of_element_located((By.XPATH, "//*[@id='room']/div[2]/div/div[2]/div[1]/div/div[3]/div/div/div[1]/div/div/div[2]/div"))).text
							p_tp=tp.split(" ")
							ws.cell(row=c+1, column=9+1).value = p_tp[0]
							print ("CHAMBRE ===")
							print (p_tp[0])
						except:
							try:
								tp = wait3.until(EC.presence_of_element_located((By.XPATH, "//*[@id='room']/div[2]/div/div[2]/div[1]/div/div[3]/div/div/div/div[1]/div/div/div[2]/div"))).text
								p_tp=tp.split(" ")
								ws.cell(row=c+1, column=9+1).value = p_tp[0]
								print ("CHAMBRE ===")
								print (p_tp[0])
							except:
								print('NO CHAMBRE')
				#VILLE
					VILLE='NO VILLE'
					run_price=extract("//a[@class='_1biqilc']/div",12,VILLE,c,YN_ville)
				#GPS
					try:
						#driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
						driver.execute_script('window.scrollBy(0,5500);')
						time.sleep(1)
						#wait4 = WebDriverWait(driver, 8)
						gps = wait3.until(EC.presence_of_element_located((By.XPATH, "//img[@alt='Carte montrant votre lieu de séjour']")))
						#"pdpPageType":1,"listingLat":46.23657,"listingLng":3.92341,"homeTier"
						a_gps=gps.get_attribute('src')
						#----------Create translation table----------
						table = str.maketrans('=&', '++')
						result_gps = a_gps.translate(table)
						split_gps=result_gps.split("+")
						#https://www.google.com/maps?ll+46.23657,3.92341+z=14&t=m&hl=fr&gl=FR&mapclient=apiv3
						coor=split_gps[1]
						long_lat=coor.split('%2C')
						#--------------Write results--------------
						print(long_lat[0])
						print(long_lat[1])
						#sheet.write(c, 12, long_lat[0])
						#sheet.write(c, 13, long_lat[1])
						ws.cell(row=c+1, column=13+1).value = long_lat[0]
						ws.cell(row=c+1, column=14+1).value = long_lat[1]
					except:
						print('NO GPS')
				#TITLE
					TITLE='NO TITLE'
					run=extract('//span[@class="_18hrqvin"]',0,TITLE,c,YN_title)
				#NAME_HOTE
					#NAME='NO NAME'
					#run=extract(X_name,2,NAME,c,YN_name)
					try:
						name = wait3.until(EC.presence_of_element_located((By.XPATH, "//div[@class='_8b6uza1']"))).text
						#a_name=name.get_attribute('aria-label')
						ws.cell(row=c+1, column=3).value = name
						print ("NAME HOTE ===")
						print (name)
					except:
						print ('NO_NAME')
				#TYPE_HOME
					TYPE='NO TYPE'
					#run=extract(X_type,7,TYPE,c,'//*[@id="room"]/div[2]/div/div[2]/div[1]/div/div[3]/div/div/div/div[3]/div/div[2]/div[1]/span')
					#//*[@id="room"]/div[2]/div/div[2]/div[1]/div/div[3]/div/div/div/div[1]/div[2]/div[1]
					try:
						#run=extract('//div[@class="_ov9erb9"]//div[@class="_1ft6jxp"]',7,TYPE,c,YN_name)
						#tp = wait3.until(EC.presence_of_element_located((By.XPATH, '//*[@id="room"]/div[2]/div/div[2]/div[1]/div/div[3]/div/div/div/div[1]/div[2]/div[1]'))).text
						tp = wait3.until(EC.presence_of_element_located((By.XPATH, "//div[@class='_504dcb']//span[@class='_1p3joamp']"))).text
						#p_tp=tp.split(".")
						ws.cell(row=c+1, column=7+1).value = tp
						print ("TYPE ===")
						print (tp)
					except:
						try:
							#tp = wait3.until(EC.presence_of_element_located((By.XPATH, '//*[@id="room"]/div[2]/div/div[2]/div[1]/div/div[3]/div/div/div/div[3]/div/div[2]/div[1]/span'))).text
							tp = wait3.until(EC.presence_of_element_located((By.XPATH, '//*[@id="room"]/div[2]/div/div[2]/div[1]/div/div[3]/div/div/div[1]/div[2]/div[1]_504dcb'))).text
							ws.cell(row=c+1, column=7+1).value = tp
							print ("TYPE ===")
							print (tp)
						except:
							print('NOTYPE')
					
				#ANCIENNETE
					try:
						#run=extract('//div[@id="host-profile"]//div[@class="_czm8crp"]',3,TYPE,c,YN_name)
						old=wait3.until(EC.presence_of_element_located((By.XPATH, '//div[@id="host-profile"]//div[@class="_czm8crp"]'))).text
						print (old)
						old1=old.split("·")
						print (old1[1])
						ws.cell(row=c+1, column=3+1).value = old1[1]
					except:
						print ('NOOLD')

				#SUPER HOTE
					try:
						a_totot = wait3.until(EC.presence_of_element_located((By.XPATH, "//div[@class='_8kd6yy']")))
						#sheet.write(c, 14, 'X')
						ws.cell(row=c+1, column=15+1).value = 'X'
						print ('SUPER HOTE')
					except:
						print('no superhote')

					if (i/10).is_integer():
						wbx.save(path_RESULT.filename)
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
						driver = webdriver.Chrome(/usr/lib/chromium-browser/chromedriver')
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
									driver = webdriver.Chrome(/usr/lib/chromium-browser/chromedriver')
									#rootdriver = webdriver.Chrome(/usr/lib/chromium-browser/chromedriver',chrome_options=chrome_options)
									driver.set_window_size(800, 800)
									wait3 = WebDriverWait(driver, 3)

			c=c+1
			i=i+1
					

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

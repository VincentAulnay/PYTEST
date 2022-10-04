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
import re

chrome_options = webdriver.ChromeOptions()
prefs = {"profile.managed_default_content_settings.images": 2}
chrome_options.add_experimental_option("prefs", prefs)
chrome_options.add_argument("window-size=2000,1000")
#chrome_options.add_argument("-headless")
#chrome_options.add_argument("-disable-gpu")
print ('▀▄▀▄▀▄ STOPBNB ▄▀▄▀▄▀')

rootdriver = webdriver.Chrome('/usr/lib/chromium-browser/chromedriver',chrome_options=chrome_options)
time.sleep(5)
#rootdriver.manage().window().setSize(1024, 768)
rootdriver.set_window_size(2000, 1000)
rootdriver.get('https://www.google.com/')
rootdriver.get('chrome://settings/')
rootdriver.execute_script('chrome.settingsPrivate.setDefaultZoom(0.5);')
rootdriver.implicitly_wait(10)
wait = WebDriverWait(rootdriver, 5)
time.sleep(5)
rootdriver.get('https://www.airbnb.fr/rooms/718608557683071399')
time.sleep(10)
try:
  
  print('version 9')
  #ele=rootdriver.find_element_by_xpath("//button[@aria-label='Avancez pour passer au mois suivant.']")
  #ele=rootdriver.find_element_by_xpath("//div[@class='_qz9x4fc']/button")
  ele = wait.until(EC.presence_of_element_located((By.XPATH, "//button[@aria-label='Avancez pour passer au mois suivant.']")))
  rootdriver.execute_script("arguments[0].scrollIntoView(true);", ele)
  rootdriver.execute_script("window.scrollBy(0,-150);")
  time.sleep(1)
  ele.click()
  time.sleep(1)
  ele.click()
  time.sleep(1)
  ele.click()
except:
  print("falde")
try:  
  time.sleep(5)
  html = rootdriver.page_source
  print('html')
  time.sleep(5)
  #soup = BeautifulSoup(html, 'html.parser')
  soup = BeautifulSoup(html, 'lxml')
  print('soup')
  time.sleep(5)
  month=soup.find('h1', attrs={"class":"_fecoyn4"}).text
  print(month)
except:
  print('falde')
try:
  month=soup.find('div', attrs={"class":"_1lds9wb"})
  print('1')
  print(month)
  the_tr= month.find_all(attrs={'aria-label':re.compile('Non'))
  print(the_tr)
  #the_tr= month.find_all('td', attrs={'aria-label':re.compile('Non')})[1]
  #the_tr= month.find_all('td', attrs={'aria-label':re.compile(r'\bNon\b')})[1]
  #div=the_tr.span.div.div.div.get_text()
  print('2')
  div=the_tr.div.get_text()
  print(div)
except:
  aaaa=1
  print('falde div')
try:
  #div=the_tr.find('div', attrs={"class": "_13m7kz7i"}).text
  intdiv=int(div)
  print(intdiv)
except:
  print('falde intdiv')
  aaaa=1

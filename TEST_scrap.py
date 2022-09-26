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
prefs = {"profile.managed_default_content_settings.images": 2}
chrome_options.add_experimental_option("prefs", prefs)
chrome_options.add_argument("window-size=2000,1000")
#chrome_options.add_argument("-headless")
#chrome_options.add_argument("-disable-gpu")
print ('▀▄▀▄▀▄ STOPBNB ▄▀▄▀▄▀')

rootdriver = webdriver.Chrome('/usr/lib/chromium-browser/chromedriver',chrome_options=chrome_options)
time.sleep(5)
#rootdriver.manage().window().setSize(1024, 768)
rootdriver.set_window_size(200, 100)
rootdriver.set_window_position(200, 50)
rootdriver.maximize_window() 
time.sleep(5)
rootdriver.set_window_size(2000, 1000)
time.sleep(5)
rootdriver.get('https://www.airbnb.fr/rooms/718608557683071399')
time.sleep(10)
try:
  ele=rootdriver.find_element_by_xpath("//button[@aria-label='Avancez pour passer au mois suivant.']")
  rootdriver.execute_script("arguments[0].scrollIntoView(true);", ele)
  rootdriver.execute_script("window.scrollBy(0,-150);")
except:
  print("falde")
try:  
  html = rootdriver.page_source
  time.sleep(1)
  soup = BeautifulSoup(html, 'html.parser')
  time.sleep(1)
  month=soup.find('h1', attrs={"class":"_fecoyn4"}).text
  print('month')
except:
  print('falde')

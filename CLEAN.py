import xlwt
import xlrd
import time
from xlutils.copy import copy
from decimal import Decimal
from decimal import *
import decimal
#import xlwings as xw
#from xlwings.constants import DeleteShiftDirection
from xlrd import open_workbook,XL_CELL_TEXT
import datetime
import datetime as dt
from tkinter import filedialog
from tkinter import *
import os
import re
import json
from urllib.request import urlopen
import pip
import pandas as pd
import openpyxl
from openpyxl import load_workbook


#--------SELECTION DU FICHIER--------
path_RESULT = Tk()
Label1 = Label(path_RESULT, text = "Select File", fg = 'red')
Label1.pack()
path_RESULT.filename =  filedialog.askopenfilename(initialdir = "/",title = "Select File",filetypes = (("Excel file","*.xlsx"),("all files","*.*")))
print (path_RESULT.filename)
NAMEFile=os.path.splitext(os.path.basename(path_RESULT.filename))[0]
print(NAMEFile)
DIR=os.path.dirname(path_RESULT.filename)
DIR2=DIR+'/'
print(os.path.dirname(path_RESULT.filename))

wb = load_workbook(path_RESULT.filename)
ws=wb.active
nrow=ws.max_row


#-------------------------------------------------------------------------------------------------------------
#-------------------SCRAP DETAILS
#-------------------

#c = ligne 2 du xls resultant

def Clean(mois):
	c=2
	while c<=nrow:
		newlist=''
		vlist=[]
		data=ws.cell(row=c, column=mois).value
		if data is not None:
			vlist=data.split(';')
			#print (len(vlist))
			i=0
			while i<len(vlist):
				if '20-6' not in vlist[i] and '21-6' not in vlist[i] and'22-6' not in vlist[i] and '23-6' not in vlist[i] and '24-6' not in vlist[i] and '25-6' not in vlist[i] and '26-6' not in vlist[i] and '30-6' not in vlist[i]:
					if newlist=='':
						newlist=newlist+vlist[i]
					else:
						newlist=newlist+';'+vlist[i]
				i=i+1
			#print('----------LIST')
			#print(vlist)
			#print('----------NEWLIST')
			#print (newlist)
			ws.cell(row=c, column=mois).value=newlist
		c=c+1


up=0
i=1
while up==0:
	V_up=ws.cell(row=1, column=i).value
	if V_up=='août 2020':
		up=1
	else:
		i=i+1
#print('DIF_Comment='+str(i))
AOUT=i

up=0
i=1
while up==0:
	V_up=ws.cell(row=1, column=i).value
	if V_up=='septembre 2020':
		up=1
	else:
		i=i+1
#print('DIF_Comment='+str(i))
SEPTEMBRE=i

up=0
i=1
while up==0:
	V_up=ws.cell(row=1, column=i).value
	if V_up=='octobre 2020':
		up=1
	else:
		i=i+1
#print('DIF_Comment='+str(i))
OCTOBRE=i


run_mars=Clean(AOUT)
run_mars=Clean(SEPTEMBRE)
run_mars=Clean(OCTOBRE)
wb.save(path_RESULT.filename)

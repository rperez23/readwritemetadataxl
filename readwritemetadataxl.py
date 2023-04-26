#! /usr/bin/python3

import os
import re
import sys
import openpyxl
import warnings
import subprocess
from openpyxl.utils.cell import get_column_letter

warnings.simplefilter(action='ignore', category=UserWarning)

numeps   = 205

startrow           = 6
innbuzzridcol      = 5
innairdatecol      = 6
inncontestantscol  = 10

#open the input xl file, exit if it fails
xlinf    = 'TPIR Drew Database.xlsx'
xlintab  = 'TPIR Drew'
try:
    workbook = openpyxl.load_workbook(filename=xlinf,data_only=True)
except:
    print("  Cannot open",xlinf)
    sys.exit(0)
if not (xlintab in workbook.sheetnames):
    print(" ",xlintab,"not in",xlinf)
    workbook.close()
ws = workbook[xlintab]



#open the output xl file, exit if it fails
xloutf   = 'zMetadata_TPIR_DREW_2023.xlsx'
xlouttab = '1. Master Metadata'
try:
    wb2 = openpyxl.load_workbook(filename=xloutf,data_only=True)
except:
    print("  Cannot open",xloutf)
    sys.exit(0)
if not (xlouttab in wb2.sheetnames):
    print(" ",xlouttab,"not in",xloutf)
    wb2.close()
ws2 = workbook[xlintab]

for i in range(0,numeps):

	row = startrow + i

	buzzrid = str(ws.cell(row,innbuzzridcol).value)
	airdate = str(ws.cell(row,innairdatecol).value)

	parts = buzzrid.split('_')

	epnum  = parts[1].replace('EP','')
	season = parts[2].replace('SR00','')
	mov    = 'ThePriceIsRight_s' + season + '_e' + epnum + '_20230410.mov'

	m = re.search('^\d+/\d+/\d+',airdate)

	if m:
		parts   = airdate.split(' ')
		airdate = parts[0]
		parts   = airdate.split('/')
		mo      = parts[0].zfill(2)
		dy      = parts[1].zfill(2)
		yr      = str(int(parts[2]) + 2000)
		airdate = yr + '-' + mo + '-' + dy
	else:
		parts = airdate.split(' ')
		airdate = parts[0] 

	print(season,':',buzzrid,':',mov,':',airdate)

workbook.close()
wb2.close()

sys.exit(0)





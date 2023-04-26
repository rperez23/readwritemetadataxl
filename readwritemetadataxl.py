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

xlinf    = 'TPIR Drew Database.xlsx'
xlintab  = 'TPIR Drew'

startrow           = 6
innbuzzridcol      = 5
innairdatecol      = 6
inncontestantscol  = 10

#open the input xl file, exit if it fails
try:
    workbook = openpyxl.load_workbook(filename=xlinf,data_only=True)
except:
    print("  Cannot open",xlinf)
    sys.exit(0)

if not (xlintab in workbook.sheetnames):
    print(" ",xlintab,"not in",xlinf)
    workbook.close()

ws = workbook[xlintab]

for i in range(0,numeps):

	row = startrow + i

	buzzrid = str(ws.cell(row,innbuzzridcol).value)
	airdate = str(ws.cell(row,innairdatecol).value)

	parts = buzzrid.split('_')

	epnum  = parts[1].replace('EP','')
	season = parts[2].replace('SR00','')
	mov    = 'ThePriceIsRight_s' + season + '_e' + epnum + '_20230410.mov'




	print(buzzrid,':',mov,':',airdate)

workbook.close()
sys.exit(0)





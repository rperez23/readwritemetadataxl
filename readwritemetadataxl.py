#! /usr/bin/python3

import os
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
    workbook = openpyxl.load_workbook(filename=xlinf)
except:
    print("  Cannot open",xlinf)
    sys.exit(0)

if not (xlintab in workbook.sheetnames):
    print(" ",xlintab,"not in",xlinf)
    workbook.close()

workbook.close()
sys.exit(0)





#! /usr/bin/python3

import re
import subprocess
import sys

inf     = 'curahee.csv'
srcdir  = 's3://s3-fremantle-uk-or-1/fremantleuk/DMS UK/FAST_CHANNEL_EDITS/Burbank/TPIR/'
outdir  = 's3://s3-fremantle-uk-or-1/fremantleuk/DMS UK/Media Files/T/ThePriceIsRight/Season'

try:
	db = open(inf,'r')
except:
	print('cannot open',inf)
	sys.exit(1)

for l in db.readlines():
	l = l.strip()
	#print(l)
	parts = l.split(',')
	mov   = parts[1]
	bzid  = parts[33]
	#print(mov,':',bzid)
	bzidmov = bzid + '.mov'

	#first see if the src mov exists

	lscmd = 'aws s3 ls "' + srcdir + bzidmov + '"'
	#print(lscmd)
	
	#TPIR_EP4853_SR0038_YR2009_DC
	m = re.search('SR00(3\d)',bzid)

	if m:
		yr = m.group(1)

	status = subprocess.run(lscmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True, shell=True) 
	
	if (status.returncode == 0):
		mvcmd = 'aws s3 mv "' + srcdir + bzidmov + '" "' + outdir + str(yr) + '/' + mov + '"'
		print(mvcmd)
		print('echo $status')
		print('echo ""')



db.close()

"""
>>> from subprocess import PIPE, run
>>> command = 'ls -l bibi'
>>> result = run(command, stdout=PIPE, stderr=PIPE, universal_newlines=True, shell=True)
>>> print(result.returncode)
1
"""

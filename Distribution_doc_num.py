#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''
Created on Jun 12, 2014

@author: Alejandro Montes Garcia
@author: Julia Efremova
@license: GPL v2
@organization: Eindhoven University of Technology
'''
import xlrd
import os.path
import string
import re
import pdb
# Path to excel file
wb = xlrd.open_workbook(os.path.join('D:/Dropbox/alejandro-julia/Database/notary_acts_19062014_copy.xlsx'))
import pylab as plb
import matplotlib as plt

sh = wb.sheet_by_index(0)

# Exclude punctuation symbols
exclude = set(string.punctuation)

text = []
# For every registry in the file the text is extracted and put into an array separating documents that are classified or need to be classified

for i in range(1, sh.nrows):  
    dummy = sh.cell(i, 9).value #Reading Datering filed from database
    if dummy:
        if type(dummy)is float:
            dummy = str(dummy)
        text.append(dummy)
    else:
        text.append('null')

print text[0:20]

text2 = dict()
for i in range(0, len(text)):
    if re.search('\d{4}', text[i])is not None:
        key = int(re.search('\d{4}', text[i]).group(0).encode('ascii'))
    else:
        key = 0
    if key>2000 or key <1300:
        key=0
    if key not in text2.keys():
        text2[key] = 1
    else:
        text2[key]= text2[key]+1
    if key=='7994':
        print text[i]
        print i
d=[]
m=[]
keylist = text2.keys()
keylist.sort()
for key in keylist:
    print "%s: %s" % (key, text2[key])
    if key>0: 
        d.append(key)
        m.append(text2[key])

plb.axis([min(d), max(d),0, max(m) ])
plb.plot(d,m,'*')
plb.xlabel('Year of Document Issue')
plb.ylabel('Number of Notary Acts')

plb.grid(True)
plb.savefig("D:/Dropbox/alejandro-julia/Distributions/Notary_acts.png")
plb.show()


   


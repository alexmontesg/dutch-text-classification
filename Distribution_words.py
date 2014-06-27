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
# Path to excel file
wb = xlrd.open_workbook(os.path.join('D:/Dropbox/alejandro-julia/Database/notary_acts_19062014_copy.xlsx'))
import pylab as plb
import pdb

shr = wb.sheet_by_index(0)
# Exclude punctuation symbols
exclude = set(string.punctuation)

text = []
documents = []

for i in range(1, shr.nrows):  
    text = shr.cell(i, 4).value
    if type(text)is float or type(text) is int:
        text = str(text)
    if shr.cell(i, 5).value:
        dummy = shr.cell(i, 5).value
        if type(dummy)is float or type(dummy) is int:
            dummy = str(dummy)
        text += dummy 
        
    if shr.cell(i, 6).value:
        dummy = shr.cell(i, 6).value
        if type(dummy)is float or type(dummy) is int:
            dummy = str(dummy)
        text += dummy 
    text = ''.join(ch for ch in text if ch not in exclude)
    #if text or type(text) is float:
    #    continue
    #else: 
    documents.append((text.lower()))

print documents[0:10]
#pdb.set_trace()
word_num = dict()

for d in documents:
    words = []
    for word in d.split():
        if(len(word) > 1):
            words.append(word)
    key = len(words)
    if key not in word_num.keys():
        word_num[key] = 1
    else:
        word_num[key]= word_num[key]+1
    
            
d=[]
m=[]
keylist = word_num.keys()
keylist.sort()
for key in keylist: #keys are size of documents, values are frequencies
    print "%s: %s" % (key, word_num[key])
    if key>0: 
        d.append(key)
        m.append(word_num[key])

plb.axis([min(d), max(d),0, max(m) ])
plb.plot(d,m,'*')
plb.xlabel('Document size (in words)')
plb.ylabel('Frequency')#

plb.grid(True)
plb.savefig("D:/Dropbox/alejandro-julia/Distributions/Words_dist.png")
plb.show()
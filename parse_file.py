'''
Created on Jun 12, 2014

@author: Alejandro Montes Garcia
@author: Julia Efremova
@license: GPL v2
@organization: Eindhoven University of Technology
'''
import xlrd
import xlsxwriter
import os.path
import re
import string
from nltk.corpus import stopwords

# Path to excel file
wbr = xlrd.open_workbook(os.path.join('/home/TUE/amontes/Dropbox/alejandro-julia/Database/notary_acts_19062014.xlsx'))
shr = wbr.sheet_by_index(0)
print "Opened notary acts. %-d rows" %shr.nrows
wbr2 = xlrd.open_workbook(os.path.join('/home/TUE/amontes/Dropbox/alejandro-julia/Database/names_to_delete.xlsx'))
shr2 = wbr2.sheet_by_index(0)
print "Opened names file. %-d rows" %shr2.nrows
wbw = xlsxwriter.Workbook('/home/TUE/amontes/Dropbox/alejandro-julia/Database/notary_acts_19062014 - classified.xlsx')
shw = wbw.add_worksheet('Data')
shw.write(0, 0, "Category")
shw.write(0, 1, "Text")
wbw2 = xlsxwriter.Workbook('/home/TUE/amontes/Dropbox/alejandro-julia/Database/notary_acts_19062014 - classified - no_stop_words_names.xlsx')
shw2 = wbw2.add_worksheet('Data')
shw2.write(0, 0, "Category")
shw2.write(0, 1, "Text")
write_row = 1

remove_names = []
exclude = set(string.punctuation)
for i in range(0, shr2.nrows):
    text = shr2.cell(i, 0).value
    if text and type(text) is not float:
        text = ''.join(ch for ch in text if ch not in exclude)
        remove_names.append(text.lower())
    text = shr2.cell(i, 1).value
    if text and type(text) is not float:
        text = ''.join(ch for ch in text if ch not in exclude)
        remove_names.append(text.lower())
    if(len(remove_names) % 10000 == 0):
        print "Added %-d names" % len(remove_names)
print "Going to delete %-d names" % len(remove_names)
remove_names = set(remove_names)
print "Going to delete %-d names" % len(remove_names)
remove_names = '|'.join(remove_names)
remove_stop = '|'.join(stopwords.words('dutch'))
remove = remove_stop + '|' + remove_names
regex = re.compile(r'\b(' + remove + r')\b', flags = re.IGNORECASE)

for i in range(1, shr.nrows):
    text = shr.cell(i, 4).value
    category = shr.cell(i, 8).value
    if shr.cell(i, 5).value:
        text += shr.cell(i, 5).value
    if shr.cell(i, 6).value:
        text += shr.cell(i, 6).value
    if not text or type(text) is float:
        continue
    if category and type(category) is not float:
        category = category.lower()
        text = ''.join(ch for ch in text if ch not in exclude)
        text = text.lower()
        shw.write(write_row, 0, category)
        shw.write(write_row, 1, text)
        text = regex.sub("", text)
        text = re.sub("\s{2,}", " ", text)
        shw2.write(write_row, 0, category)
        shw2.write(write_row, 1, text.strip())
        write_row += 1
        if(write_row % 250 == 0):
            print "%-d rows added" % write_row

wbw.close()
wbw2.close()
print "Done!"
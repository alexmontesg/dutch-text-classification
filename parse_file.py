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

# Path to excel file
wbr = xlrd.open_workbook(os.path.join('/home/TUE/amontes/Dropbox/alejandro-julia/Database/notary_acts_19062014.xlsx'))
shr = wbr.sheet_by_index(0)
wbw = xlsxwriter.Workbook('/home/TUE/amontes/Dropbox/alejandro-julia/Database/notary_acts_19062014 - classified.xlsx')
shw = wbw.add_worksheet('Data')
shw.write(0, 0, "Category")
shw.write(0, 1, "Text")
write_row = 1

# For every registry in the file the text is extracted and put into an array separating documents that are classified or need to be classified
for i in range(1, shr.nrows):
    text = shr.cell(i, 4).value
    category = shr.cell(i, 8).value
    if shr.cell(i, 5).value:
        text += shr.cell(i, 5).value
    if shr.cell(i, 6).value:
        text += shr.cell(i, 6).value
    if not text or type(text) is float:
        continue
    if category:
        shw.write(write_row, 0, category)
        shw.write(write_row, 1, text)
        write_row += 1
        if(write_row % 1000 == 0):
            print "%-d rows added" % write_row

wbw.close()
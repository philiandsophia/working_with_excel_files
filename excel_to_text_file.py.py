# -*- coding: utf-8 -*-
"""
Created on Tue Sep 12 11:38:05 2017

@author: choip
"""

import openpyxl

wb = openpyxl.load_workbook('dataset.xlsx')

sheet = wb.get_sheet_by_name('Sentiment Analysis Dataset')


print ('done')
print (wb.get_active_sheet())
alist = []


short_pos = open('positive_reddit_comments.txt','a')
short_neg = open('negative_reddit_comments.txt','a')

while len(alist) <= 10:
    for cellObj in sheet.columns:
        alist.append(cellObj)

print ('finished')



for e in alist[1]:
    for j in alist[3]:
        if e.value == 1:
            print (type(j.value),print (j.value))
            #short_pos.write(j.value + "\n")
        if e.value == 0:
            print (type(j.value),print (j.value))
            #short_neg.write(j.value + "\n")
        
short_pos.close()
short_neg.close()
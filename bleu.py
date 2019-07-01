# -*- coding: utf-8 -*-
"""
Created on Sat Jun 29 10:15:10 2019

@author: sandy16
"""
import xlwt 
from xlwt import Workbook
import pandas as pd
from nltk.tokenize import word_tokenize as tk
from nltk.translate.bleu_score import sentence_bleu


#Getting Reference data set from excel
ref = pd.read_excel('bleu reference.xlsx',usecols='A',dtype = str)
#Getting Actual data set from excel
act = pd.read_excel('bleu reference.xlsx',usecols='B',dtype = str)

#Replacing empty string for empty cells
ref.fillna('',inplace=True)
act.fillna('',inplace=True)

#Converting all data frame into list
ref_data_list = ref.values.tolist()
act_data_list = act.values.tolist()

#Initializing and calculating score
score=[]
for x in range(0,len(ref_data_list)):
    tk1 = tk(''.join(ref_data_list[x]))
    tk2 = tk(''.join(act_data_list[x]))
    if((len(tk1) == 0) and (len(tk2) == 0)):
        score.append(0)
    else:
        score.append(sentence_bleu([tk(''.join(ref_data_list[x]))],tk(''.join(act_data_list[x])) ,weights=(1,0,0,0)))



#creating workbook
wb = Workbook()
Score_sheet = wb.add_sheet("Score_sheet") 

#Bold style for Header
style = xlwt.easyxf('font: bold 1')

#Writing Headers
Score_sheet.write(0, 0, 'Reference', style) 
Score_sheet.write(0, 1, 'Actual', style) 
Score_sheet.write(0, 2, 'Score', style)


#Writing Data sets and score
row = 1 
for x in (tuple(ref_data_list)):
    Score_sheet.write(row, 0, x) 
    row += 1
row=1
for x in (tuple(act_data_list)):
    Score_sheet.write(row, 1, x)
    row += 1
row=1
for x in (tuple(score)): 
    Score_sheet.write(row, 2, x)
    row += 1
    
#saving workbook in a desired directory
wb.save("C:\\Users\\santhmoo\\Documents\\py\\bleu1.xls")



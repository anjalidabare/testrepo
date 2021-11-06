# -*- coding: utf-8 -*-
"""
Created on Thu Nov 26 17:42:37 2020

@author: Kristi - DT
"""

import os
from os import listdir
from os.path import isfile, join
import pandas as pd
import csv

currentWorkingDirectory = os.getcwd()
xlFiles = [join(currentWorkingDirectory, f) for f in listdir(currentWorkingDirectory) if (join(currentWorkingDirectory, f).endswith('xlsx'))]
Header = ['No','Field', '#SKU', 'Count', 'Trained', 'Remark']
with open('FieldMerge.csv','w') as result_file:
    wr = csv.writer(result_file, dialect='excel')
    wr.writerow(Header)
    
def readXlandAddcolumn(sourceFile,techs):
    for tech in techs:
        dataFrame = pd.DataFrame(pd.read_excel(file,sheet_name=tech))
        
            
        dataFrame.insert(1,'Tech',tech.upper())
        dataFrame.to_csv('FieldMerge.csv', mode='a', header=False, index=False)
        #dataFrame = dataFrame.iloc[1:]
        #print(dataFrame)
        



for file in xlFiles:
  
    techs= pd.ExcelFile(file).sheet_names
    readXlandAddcolumn(file,techs)
    
#print(df) 
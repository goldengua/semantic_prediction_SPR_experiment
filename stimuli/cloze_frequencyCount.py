
import pandas as pd
import numpy as np
import collections
from pandas import ExcelWriter
from pandas import ExcelFile

# open the excel file and create a dataframe

df = pd.read_excel('report.xlsx', sheet_name='Sheet2')
 
print("Column headings:")
print(df.columns)

#import packages for word processing

import nltk
import os
import jieba
import jieba.analyse
#print(df.head())
#convert into string
df=df.astype(str)

#number of questions to go through
lst = np.arange(1,56)

#create a dataframe to store the result
df2 = pd.DataFrame(lst, columns = ['number'])
wordlst = []
countlst = []

#df2 = pd.DataFrame(lst, columns = ['trial'])
#print(df2)

#for each question, go through responses
for i in lst:  
  
    select_indices = list(np.where(df["number"] == str(i))[0])
   
    response = df.iloc[select_indices,2].tolist()

    word_lst = []
    key_list = []
    
    total_response = 0
    
    for line in response:
        total_response += 1
        item = line.strip('\n\r').split('\t')
        #word segmentation
        tags = jieba.cut(item[0])
        for t in tags:
            word_lst.append(t)
            
    #count frequency and store into a dictionary        
    word_dict = {}
    with open("wordCount.txt",'w') as wf2:
        for item in word_lst:
            if item not in word_dict:
                word_dict[item] = 1
            else:
                word_dict[item] += 1
                
    sorted_by_value = sorted(word_dict.items(), key=lambda kv: kv[1],reverse = True)

    wordlst.append(sorted_by_value)
    countlst.append(total_response)
        
df2['total'] = countlst
df2['frequency'] = wordlst    

#write to excel
df2.to_csv('wordCount.txt') 







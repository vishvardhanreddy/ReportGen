'''
Created on Nov 12, 2014

@author: vvaka
'''


import numpy as numpy    
import pandas as pd

df = pd.read_csv('2.Process//nodesetup-sides.csv')

print type(df)
saved_column = df.column_name #you can also use df['column_name']
#names = df.Names

#for i in names:
#    print i


# -*- coding: utf-8 -*-
"""
Created on Wed Apr 10 14:04:27 2019

@author: eva.kung
"""

import pyodbc
import textwrap
from collections import OrderedDict
from pandas import DataFrame
import pandas as pd
import numpy as np
import xlwings as xw
print("import successfully")

cnxn = pyodbc.connect('DSN=123;DATABASE=PUBLICDB;Trusted_Connection=yes')
cursor = cnxn.cursor()

tab_discount = pd.read_excel(r"D:\eva\python\保費分布\additional table.xlsx",sheet_name = 'discount',header = 3,usecols = [7,8,9,10,11])
tab_highage = pd.read_excel(r"D:\eva\python\保費分布\additional table.xlsx",sheet_name = 'highage',header = 3,usecols = [4,5,6])
tab_crt = pd.read_excel(r"D:\eva\python\保費分布\additional table.xlsx",sheet_name = 'pro_list',header = 4,usecols = [5])

for i in range(len(tab_discount)):
    temp = cursor.execute('''insert into discount
                                    values(
                                    ?,?,?,?,?
                                    )''',
                                   tab_discount.iloc[i][0],int(tab_discount.iloc[i][1]),tab_discount.iloc[i][2],int(tab_discount.iloc[i][3]),int(tab_discount.iloc[i][4]))
    cnxn.commit()  #very important



for i in range(len(tab_crt)):
    temp = cursor.execute('''insert into pro_list
                                    values(
                                    ?
                                    )''',
                                   tab_crt.iloc[i][0])
    cnxn.commit()  #very important


for i in range(len(tab_highage)):
    temp = cursor.execute('''insert into highage
                                    values(
                                    ?,?,?
                                    )''',
                                   tab_highage.iloc[i][0],int(tab_highage.iloc[i][1]),int(tab_highage.iloc[i][2]))
    cnxn.commit()  #very important

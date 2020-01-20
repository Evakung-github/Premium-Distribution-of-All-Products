# -*- coding: utf-8 -*-
"""
Created on Tue Apr  9 22:46:04 2019

@author: eva
"""
import time
import pyodbc
import textwrap
from collections import OrderedDict
from pandas import DataFrame
import pandas as pd
import numpy as np
import xlwings as xw
print("import successfully")
pd.options.display.float_format = '{:.6%}'.format
option = 'ALL'#'ALL' #'SRCEBUS'



cnxn = pyodbc.connect('DSN=TPACTDB01_2433;DATABASE=PUBLICDB;Trusted_Connection=yes')
cursor = cnxn.cursor()

if option == 'ALL':
    temp = cursor.execute('''select distinct crtable, pay_year
                            from main
                            order by crtable''')
    par = list(iter(temp))
elif option == 'SRCEBUS':
    temp = cursor.execute('''select distinct crtable, pay_year,srcebus
                            from main
                            order by crtable''')
    par = list(iter(temp))   
else:
    print('Please check carefully.')





#pd.DataFrame(np.array(par),columns = ['crtable','pay_year','channel'])

#par = list()
#for row in temp:
#    par.append(row)

#par = [('5A34',20,'OT'),('6U55',3,'BR'),('5096',20,'AG'),('50P3',10,'BR')]
#workbook = xw.books.add()
#workbook.save('D:\eva\python\保費分布\output11.xlsx')

writer = pd.ExcelWriter(r'C:\Users\richard.hsu\Desktop\temp\output1.xlsx', engine='xlsxwriter')
wb = writer.book
format1 = wb.add_format({'num_format': '0.00%'})
format2 = wb.add_format({'num_format': '0.00'})

i = 2
for i in range(10):
    
    if len(par[i]) == 3:
        ch = par[i][2]
    else:
        ch = '\', \''.join(['AG','BR','OT','EC','WS'])
    
    
    try:
        #各通路佔比
        temp = cursor.execute('''select distinct crtable, pay_year, srcebus, sum(pl_prem) over (partition by crtable, pay_year, srcebus),
                                sum(pl_prem) over (partition by crtable) as N'各商品總和',
                                round(sum(pl_prem) over (partition by crtable, pay_year, srcebus)/sum(pl_prem) over (partition by crtable),5) as [%]
                                from main
                                where crtable =? 
                                order by crtable''',
                               par[i][0])
        t = pd.DataFrame(np.array(list(iter(temp))),columns = ['crtable','pay_year','srcebus','pl_prem','total','ratio'])
        t = t.iloc[:,1:]

        #平均保額
        temp = cursor.execute('''select crtable, pay_year, sadiscount, avg(pl_sa)
                                from dbo.main
                                where crtable ='%s' and pay_year= %d  and srcebus in ('%s')
                                group by crtable, sadiscount, pay_year
                                order by sadiscount''' % (par[i][0],par[i][1],ch))

        p = pd.DataFrame(np.array(list(iter(temp))),columns = ['crtable','pay_year','sadiscount','avg_sa'])
        p['avg_sa'] = p['avg_sa']/10000
        p = p.pivot(index = 'crtable',columns = 'sadiscount',values = 'avg_sa')
        

        cursor.commit()

        #高保額 性別占比

        temp = cursor.execute('''select distinct sadiscount,saindex, sum(pl_prem) over (partition by saindex) A, sum(pl_prem) over (partition by sadiscount),
                                sum(pl_prem) over ()
                                from main 
                                where crtable ='%s' and pay_year= %d  and srcebus in ('%s')
                                order by saindex'''% (par[i][0],par[i][1],ch))
                             
        
        qq = pd.DataFrame(np.array(list(iter(temp))),columns = ['sadiscount','saindex','pl_prem','total','T_total'])
        qq['sex'] = qq['saindex'].str[-1]
        qq.drop('saindex',axis = 1,inplace = True)
        b = pd.DataFrame(np.array([np.repeat(qq['sadiscount'].unique(),2),
                                   np.tile(np.array([0,1],dtype = str),len(qq['sadiscount'].unique()))]),index = ['sadiscount','sex']).T
        qq = pd.concat([qq,b],sort=False).fillna(0.)

        qq['sadiscount'] = qq['sadiscount'].astype(float)
        qq['pl_prem'] = qq['pl_prem'].astype(float)
        
        func = lambda x: x.sum()/float(qq['T_total'].iloc[0])
        
        qq = qq.pivot_table(index='sex', columns='sadiscount', values='pl_prem',margins = True,aggfunc = func,fill_value = 0)
        q = DataFrame(np.array([qq.iloc[-1],qq.iloc[0].div(qq.iloc[-1]),qq.iloc[1].div(qq.iloc[-1])],dtype = float),
          columns = qq.columns.values,index = ['T','M','F'])
        

        #高保額 男女 key_age 占比
        temp = cursor.execute('''select distinct saindex,key_age,sum(pl_prem) over (partition by saindex, key_age), sum(pl_prem) over (partition by saindex) B
                                from main
                                where crtable ='%s' and pay_year= %d  and srcebus in ('%s')
                                order by B'''% (par[i][0],par[i][1],ch))
                          
        rr = pd.DataFrame(np.array(list(iter(temp))),columns = ['saindex','key_age','pl_prem','total'])
        rr['ratio'] = rr['pl_prem']/rr['total']
        b['saindex'] = b['sadiscount'].apply(lambda x:'{:.2%}'.format(float(x)))+'_'+b['sex']
        b.drop(['sex','sadiscount'],axis = 1,inplace = True)
        b['key_age'] = 100
        
        rr = pd.concat([rr,b],sort=False,ignore_index = True).fillna(0)
        rr = rr.pivot_table(index='key_age', columns='saindex', values='ratio',aggfunc = np.sum)
        rr.drop(100,inplace = True)
        

        rr = rr.astype('float')


        print('商品:',par[i][0],'繳費年期:',par[i][1],'通路:',ch if len(ch) == 2 else 'All')
        #print(p)
    except:
        print('商品:',par[i][0],'繳費年期:',par[i][1],'通路:',ch if len(ch) == 2 else 'All','   error')
    
    else:
        name = str(par[i][0])+'_'+str(par[i][1])+'_'+(ch if len(ch) == 2 else 'All')
        
#        try:
#            sheets = wb.add_worksheet(name)
#        except:
#            sheets = writer.sheets[name]
#
#        sheets.set_column('B:I', 18, format1)
        #sheets.cells(2,'B').value = p
        p.to_excel(writer,sheet_name = name,startcol = 2,startrow = 2)
        #sheets.cells(6,'B').value = q
        q.to_excel(writer,sheet_name = name,startcol = 2,startrow = 6)
        #sheets.cells(17,'B').value = rr
        rr.to_excel(writer,sheet_name = name,startcol = 2,startrow = 17)
        #sheets.cells(2,'I').value = t
        t.to_excel(writer,sheet_name = name,startcol = 9,startrow = 2)
        sheets = writer.sheets[name]
        sheets.set_column('B:Z', 18, format1)
        sheets.set_column('O:O', 18, format1)
        sheets.set_column('M:N', 18, format2)
        cursor.commit()
    
writer.save()
writer.close()
cursor.close()

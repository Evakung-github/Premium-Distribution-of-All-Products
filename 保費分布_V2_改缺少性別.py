# -*- coding: utf-8 -*-
"""
Created on Tue Apr  9 22:46:04 2019

@author: eva
"""

import pyodbc
import textwrap
from collections import OrderedDict
from pandas import DataFrame
import pandas as pd
import numpy as np
import xlwings as xw
print("import successfully")
pd.options.display.float_format = '{:.2%}'.format

cnxn = pyodbc.connect('DSN=TPACTDB01_2433;DATABASE=PUBLICDB;Trusted_Connection=yes')
cursor = cnxn.cursor()


temp = cursor.execute('''select distinct crtable, pay_year, srcebus
                            from main
                            order by crtable''')
par = list()
for row in temp:
    par.append(row)

#par = [('5A34',20,'OT'),('6U55',3,'BR'),('5096',20,'AG'),('50P3',10,'BR')]
workbook = xw.books.add()
workbook.save('D:\eva\python\保費分布\output5.xlsx')
i = 2
for i in range(len(par)):
    #平均保額
    
    try:
        temp = cursor.execute('''select distinct crtable, pay_year, srcebus, sum(pl_prem) over (partition by crtable, pay_year, srcebus),
                                sum(pl_prem) over (partition by crtable) as N'各商品總和',
                                round(sum(pl_prem) over (partition by crtable, pay_year, srcebus)/sum(pl_prem) over (partition by crtable),5) as [%]
                                from main
                                where crtable =? 
                                order by crtable''',
                               par[i][0])
        
        crtable = list()
        pay_year = list()
        srcebus = list()
        pl_prem = list()
        total = list()
        ratio = list()
        for row in temp:
            pay_year.append(row[1])
            srcebus.append(row[2])
            pl_prem.append(row[3])
            total.append(row[4])
            ratio.append('{:.2%}'.format(row[5]))
        
        table = OrderedDict((
        ("pay_year",pay_year),
        ("srcebus",srcebus),
        ("pl_prem",pl_prem),
        ("total",total),   
        ("ratio",ratio)
        ))

        t=DataFrame(table)    
        
 
        temp = cursor.execute('''select crtable, pay_year, sadiscount, avg(pl_sa)
                                from dbo.main
                                where crtable =? and pay_year=? and srcebus =?
                                group by crtable, sadiscount, pay_year
                                order by sadiscount''',
                              par[i][0],par[i][1],par[i][2])


        discount = list()
        avg_sa = list()
        crtable = list()

        for row in temp:
            discount.append(row[2])
            avg_sa.append(float(row[3]/10000))
            crtable.append(row[0])

        table = OrderedDict((
        ("crtable",crtable),
        ("discount",discount),
        ("avg_sa",avg_sa)
        ))

        d = DataFrame(table)   

        p = d.pivot(index='crtable', columns='discount', values='avg_sa')
        
        #print(p)
        cursor.commit()

        #高保額 性別占比

        temp = cursor.execute('''select distinct sadiscount,saindex, sum(pl_prem) over (partition by saindex) A, sum(pl_prem) over (partition by sadiscount),
                                sum(pl_prem) over ()
                                from main 
                                where crtable =? and pay_year=? and srcebus =?
                                order by saindex''',
                             par[i][0],par[i][1],par[i][2])
        discount=list()
        sex=list()
        pl_prem=list()
        total =list()
        ratio=list()
        T_total = list()
        n =0
        for row in temp:
            discount.append(row[0])
            sex.append(row[1][-1]) #類似right用法
            pl_prem.append(int(row[2]))
            total.append(int(row[3]))
            T_total.append(row[4])
     
        table = OrderedDict((
                  ("discount",discount),
                  ("sex",sex),
                  ("pl_prem",pl_prem),
                  ("total",total),
                  ("T_total",T_total)
                  ))

        d = DataFrame(table) 
        func = lambda x: x.sum()/float(T_total[0])
        #float_formatter = lambda x: "{:.2%}".format(x)
        #np.set_printoptions(formatter={'float_kind':float_formatter})
        b = pd.DataFrame(np.array([np.repeat(d['discount'].unique(),2),np.tile(np.array([0,1],dtype = str),len(d['discount'].unique()))]),index = ['discount','sex']).T

        q =  pd.concat([d,b],sort = False)
        q = q.fillna(0)
        q['discount'] = q['discount'].astype(float)
        q1 = q[d.columns]

        q = q1.pivot_table(index='sex', columns='discount', values='pl_prem',margins = True,aggfunc = func,fill_value = 0)
        q = DataFrame(np.array([q.iloc[-1],q.iloc[0].div(q.iloc[-1]),q.iloc[1].div(q.iloc[-1])],dtype = float),
          columns = q.columns.values,index = ['T','M','F'])
        
        #print(p)
        #print('\n')

        #高保額 男女 key_age 占比
        temp = cursor.execute('''select distinct saindex,key_age,sum(pl_prem) over (partition by saindex, key_age), sum(pl_prem) over (partition by saindex) B
                                from main
                                where crtable =? and pay_year=? and srcebus =?
                                order by B''',
                             par[i][0],par[i][1],par[i][2])

        saindex=list()
        key_age=list()
        pl_prem=list()
        total = list()
        ratio=list()

        for row in temp:
            saindex.append(row[0])
            key_age.append(row[1])
            pl_prem.append(int(row[2]))
            total.append(int(row[3]))
            ratio.append('{:.2%}'.format(int(row[2])/int(row[3])))

        table = OrderedDict((
        ("sa_index",saindex),
        ("key_age",key_age),
        ("pl_prem",pl_prem),
        ("ratio",ratio)
        ))

        b['sa_index'] = b['discount'].apply(lambda x:'{:.2%}'.format(float(x)))+'_'+b['sex']
        b['key_age'] = 100
        d = DataFrame(table)
        column_names = d.columns
        d = pd.concat([d,b[['sa_index','key_age']]],sort = False,ignore_index = True).fillna(0)
        d[['key_age','pl_prem']]= d[['key_age','pl_prem']].astype('int')
        d = d[column_names]

        r = d.pivot_table(index='key_age', columns='sa_index', values='ratio',aggfunc = np.sum)
        #r.index = r.index.droplevel(level = 1)
        r = r.drop(100)

        print('商品:',par[i][0],'繳費年期:',par[i][1],'通路:',par[i][2])
        #print(p)
    except:
        print('商品:',par[i][0],'繳費年期:',par[i][1],'通路:',par[i][2],'   error')
    
    else:
        name = str(par[i][0])+'_'+str(par[i][1])+'_'+par[i][2]
        
        try:
            sheets = workbook.sheets.add(name)
        except:
            sheets = workbook.sheets(name)
        sheets.cells(2,'B').value = p
        sheets.cells(6,'B').value = q
        sheets.cells(17,'B').value = r
        sheets.cells(2,'I').value = t
        cursor.commit()
    
workbook.save()
workbook.close()
cursor.close()
par[0]
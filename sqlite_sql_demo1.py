# -*- coding: utf-8 -*-
"""
Created on Tue Nov 28 08:13:07 2017

@author: lenovo
"""

import sqlite3
import pandas as pd
import numpy as np

def createDataBase():
    cn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')

    cn.execute('''CREATE TABLE IF NOT EXISTS PE_PB
           (GPDM TEXT,
           RQ TEXT,
           PE_LYR REAL,
           PE_TTM REAL,
           PB REAL);''')

    cn.execute('''CREATE UNIQUE INDEX IF NOT EXISTS GPDM_RQ_PEPB ON PE_PB(GPDM,RQ);''')

    cn.execute('''CREATE TABLE IF NOT EXISTS GPDMB
           (GPDM TEXT PRIMARY KEY,
           GPMC TEXT,
           SSRQ TEXT);''')

dbcn = sqlite3.connect(r'd:\hyb\stockdata.db')
curs = dbcn.cursor()


curs.execute('''select code,date,pe_lyr,pe_ttm,pb from pe ;''')
data = curs.fetchall()

df=pd.DataFrame(data,columns=['gpdm','rq','pe_lyr','pe_ttm','pb'])
df['gpdm']=df['gpdm'].map(lambda x:x[:6]+('.SH' if x[0]=='6' else '.SZ'))

dt=np.array(df).tolist()

createDataBase()
dbcn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')
dbcn.executemany('INSERT OR IGNORE INTO PE_PB (GPDM,RQ,pe_LYR,PE_TTM,PB) VALUES (?,?,?,?,?)', dt)
dbcn.commit()
dbcn.close()



##保留最新户数
#df1=df.drop_duplicates(['gpdm'],keep='first')
#
##curs.execute('''select gpdm,rq,gdhs from gdhs 
##          where rq=='2017-09-30';''')
#curs.execute('''select gpdm,rq,gdhs from gdhs 
#          where rq=='2016-09-30';''')
#data = curs.fetchall()
#df2=pd.DataFrame(data,columns=['gpdm','rq','gdhs'])
#
#df3=pd.merge(df1, df2, on='gpdm')
#df3['zjbl']=df3['gdhs_x']/df3['gdhs_y']-1
#df4=df3.loc[df3['zjbl']<-0.1]
#df4=df4.sort_values(by="zjbl")



# select
#curs.execute('''select gpdm,rq,gdhs from gdhs 
#          where rq>='2017-09-30' order by rq desc;''')
#data = curs.fetchall()
#df=pd.DataFrame(data,columns=['gpdm','rq','gdhs'])
##保留最新户数
#df1=df.drop_duplicates(['gpdm'],keep='first')
#
##curs.execute('''select gpdm,rq,gdhs from gdhs 
##          where rq=='2017-09-30';''')
#curs.execute('''select gpdm,rq,gdhs from gdhs 
#          where rq=='2016-09-30';''')
#data = curs.fetchall()
#df2=pd.DataFrame(data,columns=['gpdm','rq','gdhs'])
#
#df3=pd.merge(df1, df2, on='gpdm')
#df3['zjbl']=df3['gdhs_x']/df3['gdhs_y']-1
#df4=df3.loc[df3['zjbl']<-0.1]
#df4=df4.sort_values(by="zjbl")


def createDataBase():
    cn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')

    cn.execute('''CREATE TABLE IF NOT EXISTS PE_PB
           (GPDM TEXT,
           RQ TEXT,
           PE_LYR REAL,
           PE_TTM REAL,
           PB REAL);''')

    cn.execute('''CREATE UNIQUE INDEX IF NOT EXISTS PE_PB ON PE_PB(GPDM,RQ);''')

    cn.execute('''CREATE TABLE IF NOT EXISTS GPDMB
           (GPDM TEXT PRIMARY KEY,
           GPMC TEXT,
           SSRQ TEXT);''')
    


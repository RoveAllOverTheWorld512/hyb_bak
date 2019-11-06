# -*- coding: utf-8 -*-
"""
Created on Mon Nov 27 09:40:37 2017

@author: lenovo
"""
import sqlite3
import time
import datetime
import xlwings as xw
import pandas as pd
import numpy as np

########################################################################
#建立数据库
########################################################################
def createDataBase():
    cn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')

    cn.execute('''CREATE TABLE IF NOT EXISTS GDHS
           (GPDM TEXT NOT NULL,
           RQ TEXT NOT NULL,
           GDHS INTEGER NOT NULL);''')
    cn.execute('''CREATE UNIQUE INDEX IF NOT EXISTS GPDM_RQ ON GDHS(GPDM,RQ);''')

##########################################################################
#判断3种类型字符串日期"20170101"、"2017-01-01"、"2017/01/01"是否有效
'''
    strptime(string, format) -> struct_time
    
    Parse a string to a time tuple according to a format specification.
    See the library reference manual for formatting codes (same as
    strftime()).
    
    Commonly used format codes:
    
    %Y  Year with century as a decimal number.
    %m  Month as a decimal number [01,12].
    %d  Day of the month as a decimal number [01,31].
    %H  Hour (24-hour clock) as a decimal number [00,23].
    %M  Minute as a decimal number [00,59].
    %S  Second as a decimal number [00,61].
    %z  Time zone offset from UTC.
    %a  Locale's abbreviated weekday name.
    %A  Locale's full weekday name.
    %b  Locale's abbreviated month name.
    %B  Locale's full month name.
    %c  Locale's appropriate date and time representation.
    %I  Hour (12-hour clock) as a decimal number [01,12].
    %p  Locale's equivalent of either AM or PM.
'''
#注意：年必须4位，月、日必须2位
##########################################################################
def isVaildDate(date):
    if "-" in date:
        try :
            if ":" in date:
                time.strptime(date, "%Y-%m-%d %H:%M:%S")
            else:
                time.strptime(date, "%Y-%m-%d")
            return True
        except:
            return False 

    if "/" in date:
        try:
            time.strptime(date, "%Y/%m/%d")
            return True
        except:
            return False
    #注意长度8位的情况有可能是"2017/9/3"或"2017-9-3" 所以放在后面           
    if len(date)==8 :
        try:
            time.strptime(date, "%Y%m%d")
            return True
        except:
            return False

    return False
        

if __name__ == "__main__":  

    app=xw.App(visible=True,add_book=False)
    app.display_alerts=False
    app.screen_updating=False
    filepath=r'd:\hyb\股东户数_20171127.xlsx'
    wb=app.books.open(filepath)

    sht1 = wb.sheets.add(name='gdhs')
    sht2 = wb.sheets['最新股东户数']
    
    rng2 = sht2.range('A1').expand('down') 
    rng1 = sht1.range('A1')
    data=rng2.options(ndim=2).value
    for i in range(1,len(data)):
        dt=data[i][0]
        data[i][0]= dt + ('.SH' if dt[0]=='6' else '.SZ')
    rng1.value=data
    
    
    my_date_handler = lambda year, month, day, **kwargs: "%04i-%02i-%02i" % (year, month, day)
    rng2 = sht2.range('I1').expand('down') 
    rng1 = sht1.range('B1')
    data=rng2.options(ndim=2,dates=my_date_handler).value
    for i in range(1,len(data)):
        data[i][0]=data[i][0].replace('/','-')
    rng1.value=data
    rng1 = sht1.range('B1').expand('down') 
    rng1.number_format="yyyy-mm-dd;@"
    
#    rng2 = sht2.range('D1').expand('down') 
#    rng1 = sht1.range('C1')
#    rng1.value=rng2.options(ndim=2).value
#    
#    
#    rng1 = sht1.range('A2').expand('table') 
#    data=rng1.value
#    wb.save()
#    wb.close()
    
    app.screen_updating=True
    app.display_alerts=True
#    app.quit()
    
    df=pd.DataFrame(data,columns=['gpdm','rq','gdhs'])
    #检查日期是否有效，去掉无效日期行
#    df=df[df['rq'].map(isVaildDate)]
    #转换日期格式
#    df['rq']=df['rq'].map(lambda x:x.replace('/','-'))
    #重新排列列顺序

    #将pandas重新转换为list
    dt=np.array(df).tolist()
    #导入数据库

#    createDataBase()
#    dbcn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')
#    dbcn.executemany('INSERT OR IGNORE INTO GDHS (GPDM,RQ,GDHS) VALUES (?,?,?)', dt)
#    dbcn.commit()
#    dbcn.close()
    
    


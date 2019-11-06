# -*- coding: utf-8 -*-
"""
Created on Tue Dec  5 16:28:50 2017

@author: lenovo
"""
import sqlite3

def createDataBase():
    cn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')
    '''
    GJFQ股价复权控制表
    
    GPDM股票代码
    FHPGRQ最近分红配股日期
    GJFQRQ最后计算股价复权日期
    
    GPDM为主键
    '''
    cn.execute('''CREATE TABLE IF NOT EXISTS GJFQ
           (GPDM TEXT PRIMARY KEY NOT NULL,
            FHPGRQ TEXT,
            GJFQRQ TEXT);''')


if __name__ == "__main__":  
    createDataBase()
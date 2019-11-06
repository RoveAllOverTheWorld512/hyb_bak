# -*- coding: utf-8 -*-
"""
Created on Wed Nov 15 20:36:33 2017

@author: lenovo
"""

import sqlite3
import pandas as pd
import os
import sys
import re

##########################################################################
#读取当前工作路径盘符
##########################################################################
def getdisk():
    return sys.argv[0][:2]

######################################################################################
#检测路径是否存在，不存则创建
######################################################################################    
def exsit_path(pth):
    if not os.path.exists(pth) :
        os.makedirs(pth)

########################################################################
#股票代码表
########################################################################
def gpdmdict():
    fn = getdisk()+'\\hyb\\gpdmb.txt'
    with open(fn) as f:
        gpdmb = f.read()
        f.close()

    dmb = re.findall('(\d{6})\t(.+)\n',gpdmb)
    dm = {}
    for (gpdm,gpmc) in dmb :
        dm[gpdm] = gpmc

    return dm


########################################################################

if __name__ == '__main__':

    gpdm="002496"
    gpmc = gpdmdict()[gpdm]
    pth =  'D:/公司研究/'+gpmc
     
    exsit_path(pth)
    
    fn = pth+'/'+gpdm+gpmc+'市盈率.xlsx'
    
    
    conn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')
    
    sql="select date,pe_lyr,pe_ttm from pe where code='" +gpdm+"';"
    df=pd.read_sql_query(sql, con=conn)
#    df=pd.read_sql_query(sql, con=conn,parse_dates=['DATE'])
#    df=df.set_index('DATE')
    conn.close()
    
    writer = pd.ExcelWriter(fn, engine='xlsxwriter')
    
    df.to_excel(writer, sheet_name='市盈率',index=False)
    
    workbook = writer.book
    worksheet = writer.sheets['市盈率']
    
    format2 = workbook.add_format({'num_format': '0.00'})
    format3 = workbook.add_format({'num_format': 'yyyy-mm-dd'})
    
    worksheet.set_column('A:A', 11, format3)
    worksheet.set_column('B:C', 8, format2)
    
    
    writer.save()
    
    print("文件保存为："+fn)

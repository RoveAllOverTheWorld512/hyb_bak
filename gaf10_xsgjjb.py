# -*- coding: utf-8 -*-
"""
Created on Tue Feb  7 22:25:24 2017
限售股解禁
@author: lenovo
"""
from pyquery import PyQuery as pq
import pandas as pd
import re
import os
import sys

########################################################################
#检测是不是可以转换成浮点数
########################################################################
def str2float(num):
    try:
        return float(num)
    except ValueError:
        return num

##########################################################################
#读取当前工作路径盘符
##########################################################################
def getdisk():
    return sys.argv[0][:2]

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

######################################################################################
#检测路径是否存在，不存则创建
######################################################################################    
def exsit_path(pth):
    if not os.path.exists(pth) :
        os.makedirs(pth)

########################################################################
#提取历年股本变动数据
########################################################################
def lngbbd(gpdm): 
    
    gpmc = gpdmdict()[gpdm]
    data=[]
    data.append(['股票代码','股票名称','变更日期','总股本(万股)','流通A股(万股)','实际流通A股(万股)','变更原因'])

    url  = "http://web-f10.gaotime.com/stock/"+gpdm+"/gbjg/lngbbd.html"

    html = pq(url,encoding="utf-8")

    tb = html('tr')

    for i in range(1,len(tb)-1) :
        row=pq(tb.eq(i).html())
        tc=row('td')
        rowdat=[gpdm,gpmc]
        for j in range(0,len(tc)):
            col=row.find('td').eq(j).text()
            rowdat.append(str2float(col))

        data.append(rowdat)
 
    return pd.DataFrame(data[1:],columns=data[0])       


def xsjj(gpdm):
    data = []
    data.append(['股票代码','股东名称','流通日期','新增可售股数(万股)','限售类型','限售条件说明'])
    url  = "http://web-f10.gaotime.com/stock/"+gpdm+"/gbjg/xsjj.html"

    html = pq(url,encoding="utf-8")

    tb = html('tr')

    for i in range(1,len(tb)-1) :
        row=pq(tb.eq(i).html())
        tc=row('td')
        rowdat=[gpdm]
        if len(tc)==5 :
            gdmc=row.find('td').eq(0).text()
        else:
            rowdat.append(gdmc)
        for j in range(0,len(tc)):
            col=row.find('td').eq(j).text()
            rowdat.append(str2float(col))

        data.append(rowdat)

    return pd.DataFrame(data[1:],columns=data[0])       

    
if __name__ == "__main__":  
    gpdm = "002496"
    gpmc = gpdmdict()[gpdm]

    pth =  'D:/公司研究/'+gpmc
    exsit_path(pth)
    fn = pth+'\\'+gpdm+gpmc+'限售股解禁时间表.xlsx'
    if os.path.exists(fn):
        os.remove(fn)

    df1 = lngbbd(gpdm)
    df2 = xsjj(gpdm)
    df3 = df2.groupby(['流通日期']).sum()
    df3['流通日期']=df3.index
    df3=df3.loc[:,['流通日期','新增可售股数(万股)']]
    writer=pd.ExcelWriter(fn,engine='xlsxwriter')

    df1.to_excel(writer, sheet_name='历年股本变动',index=False)
    df2.to_excel(writer, sheet_name='限售股解禁',index=False)   
    df3.to_excel(writer, sheet_name='限售股解禁汇总',index=False)   

    workbook = writer.book
    worksheet1 = writer.sheets['历年股本变动']
    worksheet2 = writer.sheets['限售股解禁']
    worksheet3 = writer.sheets['限售股解禁汇总']

    format1 = workbook.add_format({'num_format': '0.0000'})
    format2 = workbook.add_format({'num_format': '0.00'})
    format3 = workbook.add_format({'num_format': 'yyyy-mm-dd'})
    format4 = workbook.add_format({'align': 'center'})
    
    worksheet1.set_column('A:B', 10, format4)
    worksheet1.set_column('C:C', 12, format3)
    worksheet1.set_column('D:F', 17, format2)
    worksheet1.set_column('G:G', 35)

    worksheet2.set_column('A:A', 10, format4)
    worksheet2.set_column('B:B', 38)
    worksheet2.set_column('C:C', 11, format3)
    worksheet2.set_column('D:D', 20, format2)
    worksheet2.set_column('E:E', 20)
    worksheet2.set_column('F:F', 13, format4)
    
    worksheet3.set_column('A:A', 10, format4)
    worksheet3.set_column('B:B', 21, format2)
         
    writer.save()

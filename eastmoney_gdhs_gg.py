# -*- coding: utf-8 -*-
"""
Created on Tue Feb  7 22:25:24 2017

@author: lenovo
"""


import os
import sys
from configobj import ConfigObj
from selenium import webdriver
import pandas as pd
import re
import winreg
import datetime
import time

########################################################################
#初始化本程序配置文件
########################################################################
def iniconfig():
    myname=filename(sys.argv[0])
    wkdir = os.getcwd()
    inifile = os.path.join(wkdir,myname+'.ini')  #设置缺省配置文件
    return ConfigObj(inifile,encoding='GBK')

#########################################################################
#读取键值
#########################################################################
def readkey(config,key):
    keys = config.keys()
    if keys.count(key) :
        return config[key]
    else :
        return ""

#########################################################################
#读取文件名
#########################################################################
def filename(pathname):
    wjm = os.path.splitext(os.path.basename(pathname))
    return wjm[0]


########################################################################
#检测是不是可以转换成浮点数
########################################################################
def str2float(num):
    try:
        return float(num)
    except ValueError:
        return num

def exsit_path(pth):
    if not os.path.exists(pth) :
        os.makedirs(pth)

def getgdhs(gpdm):
   
#    browser = webdriver.Firefox()
    browser = webdriver.PhantomJS()
    browser.maximize_window()
    today = datetime.datetime.now().strftime("%Y/%m/%d") 
    data = []
    fld1=['股票代码','股票简称','截止日期','区间涨幅(%)','本次户数','上次户数','增减户数','增加比例','户均市值(万)','户均持股(万)',
                 '总市值(亿)','总股本(亿)','股本变动','股本变动原因','公告日期']
    fld2=['gpdm','gpmc','jzrq','qjzf','bchs','schs','zjhs','zjbl','hjsz','fjcg',
                 'zsz','zgb','gbbd','gbbdyy','gbrq']
    url = "http://data.eastmoney.com/gdhs/detail/"+gpdm+".html"
    
    browser.get(url)
    
    gpmc=browser.find_element_by_class_name("tit")
    gpmc=gpmc.text[0:gpmc.text.find(gpdm)-1].replace(" ","").replace("*","")

    pgnv = browser.find_elements_by_id("PageCont")
    pgs=pgnv[0].find_elements_by_tag_name("a")
    if len(pgs)==0 :
        pg=1
    else:
        pg=int(pgs[len(pgs)-3].text)
        
    tbl = browser.find_elements_by_id("dt_1")
    tbody = tbl[0].find_elements_by_tag_name("tbody")
    tblrows = tbody[0].find_elements_by_tag_name('tr')
       
    for j in range(len(tblrows)):
        rowdat = [gpdm,gpmc]
        tblcols = tblrows[j].find_elements_by_tag_name('td')
        for i in range(len(tblcols)):
            coldat = str2float(tblcols[i].text)
            rowdat.append(coldat)
    
        data.append(rowdat)
        
    if pg>1 :
        for k in range(2,pg+1):
            elem = browser.find_element_by_id("PageContgopage")
            elem.clear()
            elem.send_keys(k)
            elem = browser.find_element_by_class_name("btn_link")        
            elem.click()
            time.sleep(2)
            tbl = browser.find_elements_by_id("dt_1")
            tbody = tbl[0].find_elements_by_tag_name("tbody")
            tblrows = tbody[0].find_elements_by_tag_name('tr')
               
            for j in range(len(tblrows)):
                rowdat = [gpdm,gpmc]
                tblcols = tblrows[j].find_elements_by_tag_name('td')
                for i in range(len(tblcols)):
                    coldat = str2float(tblcols[i].text)
                    rowdat.append(coldat)
            
                data.append(rowdat)
 
    browser.quit()
  
    df = pd.DataFrame(data,columns=fld2)
    
    df.sort_values(by="jzrq", ascending=False, inplace=True)    #排序，逆序
    df1=df.loc[0:0]
    df2=df1.append(df, ignore_index=True)
    df2.loc[0,'jzrq']=today
    df2.loc[0,'qjzf']=None
    df2.loc[0,'zsz']=None
    df2.loc[0,'gbbd']=None
    df2.loc[0,'gbbdyy']=None
    df2.loc[0,'gbrq']=None
    df2.sort_values(by="jzrq", ascending=True, inplace=True)    #排序
        
    df2.columns=fld1
    
    pth = 'd:\\公司研究\\'+gpmc
    
    exsit_path(pth)
    
    fn = pth+'\\'+gpdm+gpmc+'股东户数.xlsx'
    
    writer = pd.ExcelWriter(fn, engine='xlsxwriter')

    df2.to_excel(writer, sheet_name='股东户数',index=False)

    workbook = writer.book
    worksheet = writer.sheets['股东户数']

    format1 = workbook.add_format({'num_format': '#,##0.00'})
    format2 = workbook.add_format({'num_format': '0'})
    
    worksheet.set_column('D:L', 12, format1)
    worksheet.set_column('E:G', 12, format2)
    writer.save()

########################################################################
#获取本机通达信安装目录，生成自定义板块保存目录
########################################################################
def gettdxblkdir():
    try :
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\华西证券华彩人生")
        value, type = winreg.QueryValueEx(key, "InstallLocation")
        return value + '\\T0002\\blocknew'
    except :
        print("本机未安装【华西证券华彩人生】软件系统。")
        sys.exit()

#############################################################################
#股票列表,通达信板块文件调用时wjtype="tdxbk"
#############################################################################
def zxglist(zxgfn,wjtype=""):
    zxglst = []
    p = "(\d{6})"
    if wjtype == "tdxblk" :
        p ="\d(\d{6})"
    if os.path.exists(zxgfn) :
        #用二进制方式打开再转成字符串，可以避免直接打开转换出错
        with open(zxgfn,'rb') as dtf:
            zxg = dtf.read()
            if zxg[:3] == b'\xef\xbb\xbf' :
                zxg = zxg.decode('UTF8','ignore')   #UTF-8
            elif zxg[:2] == b'\xfe\xff' :
                zxg = zxg.decode('UTF-16','ignore')  #Unicode big endian
            elif zxg[:2] == b'\xff\xfe' :
                zxg = zxg.decode('UTF-16','ignore')  #Unicode
            else :
                zxg = zxg.decode('GBK','ignore')      #ansi编码
        zxglst =re.findall(p,zxg)
    else:
        print("文件%s不存在！" % zxgfn)
    if len(zxglst)==0:
        print("股票列表为空,请检查%s文件。" % zxgfn)

    zxg = list(set(zxglst))
    zxg.sort(key=zxglst.index)

    return zxg

if __name__ == "__main__":  

    zxgfile="zxg.blk"
    tdxblkdir = gettdxblkdir()
    zxgfile = os.path.join(tdxblkdir,zxgfile)
    zxglb = zxglist(zxgfile,"tdxblk")
    j=155       #最小值为1
    for i in range(j-1,len(zxglb)):
        gpdm=zxglb[i]
        print('共有%d只股票,正在处理第%d只股票，请等待。' %(len(zxglb),i+1))
        getgdhs(gpdm)


    
        

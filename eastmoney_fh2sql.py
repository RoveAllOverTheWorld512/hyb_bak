# -*- coding: utf-8 -*-
"""
本程序将从东方财富网抓取股东户数信息存入sqlite数据库
"""

from selenium import webdriver
import sqlite3
import time
import datetime
import re

def createDataBase():
    cn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')
    '''
    GPDM股票代码
    RQ股权登记日，除权除息日一般为该股下一个交易日
    FH每股分红
    SZG每股送转股
    PGBL每股配股比例
    PGJ每股配股价
    GPDM与RQ构成为唯一索引
    '''
    cn.execute('''CREATE TABLE IF NOT EXISTS FH
           (GPDM TEXT NOT NULL,
           RQ TEXT NOT NULL,
           FH REAL NOT NULL DEFAULT 0.00,
           SZG REAL NOT NULL DEFAULT 0.00,
           PGBL REAL NOT NULL DEFAULT 0.00,
           PGJ REAL NOT NULL DEFAULT 0.00);''')
    
    cn.execute('''CREATE UNIQUE INDEX IF NOT EXISTS GPDM_RQ_FH ON FH(GPDM,RQ);''')


#########################################################################################
#提取网页数
#########################################################################################   
def getpgs(bgq):
    url = "http://data.eastmoney.com/yjfp/" + bgq + ".html"
    print('正在处理报告期：%s' % bgq)
    browser = webdriver.PhantomJS()
    browser.get(url)
    time.sleep(5)
    
    pgnv = browser.find_elements_by_id("PageCont")
    pgs=pgnv[0].find_elements_by_tag_name("a")
    if len(pgs)==0 :
        pg=1
    else:
        if len(pgs)>=4 and len(pgs)<=7:
            pg=len(pgs)-2
        else:
            if pgs[len(pgs)-3].text=='...' :
                elem = pgs[len(pgs)-3]
                elem.click()
                time.sleep(3)
                pgnv = browser.find_elements_by_id("PageCont")
                pgs=pgnv[0].find_elements_by_tag_name("a")
                pg=int(pgs[len(pgs)-3].text)+1
            else:
                pg=int(pgs[len(pgs)-3].text)

    browser.quit()

    return pg           

#########################################################################################
#提取数据，bgq报告期、pgn页码、pgs总页数，成功返回True，不成功返回False
#########################################################################################   
def getdata(bgq,pgn,pgs):
    f = open("fhdataerr.log", "a")
    f.write("报告期："+bgq+"共有"+str(pgs)+"页，正在处理第"+str(pgn)+"页，请等待。"+"\r\n")
    dbcn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')
    url = "http://data.eastmoney.com/yjfp/" + bgq + ".html"
    browser = webdriver.PhantomJS()
    browser.get(url)
    time.sleep(5)
    print('报告期：%s共有%d页，正在处理第%d页，请等待。' %(bgq,pgs,pgn))
    if pgs>1 :
        try :
            elem = browser.find_element_by_id("gopage")
            elem.clear()
            elem.send_keys(pgn)
            elem = browser.find_element_by_class_name("btn_link")        
            elem.click()
            time.sleep(5)
        except:
            f.close()
            dbcn.close()
            return False
        
            
    try :
        tbl = browser.find_elements_by_id("dt_1")
        tbody = tbl[0].find_elements_by_tag_name("tbody")
        tblrows = tbody[0].find_elements_by_tag_name('tr')
    except :
        f.close()
        dbcn.close()
        return False
 
    data = []
      
    for j in range(len(tblrows)):
        try:
            tblcols = tblrows[j].find_elements_by_tag_name('td')
            dm=tblcols[0].text
            sgstr=tblcols[4].text
            zgstr=tblcols[5].text
            fhstr=tblcols[6].text
            rq = tblcols[15].find_elements_by_tag_name('span')[0].get_property("title")        
        except:
            f.close()
            dbcn.close()
            return False
           
            
        dm=dm+('.SH' if dm[0]=='6' else '.SZ')

        sgs = re.findall ('([\d\.]+)送([\d\.]+)',sgstr)
        if len(sgs) == 1:
            sg = float(sgs[0][1])/float(sgs[0][0])
        else :
            sg=0

        zgs = re.findall ('([\d\.]+)转([\d\.]+)',zgstr)
        if len(zgs) == 1:
            zg = float(zgs[0][1])/float(zgs[0][0])
        else :
            zg=0
            
        szg=sg+zg

        fhs = re.findall ('([\d\.]+)派([\d\.]+)',fhstr)
        if len(fhs) == 1:
            fh = float(fhs[0][1])/float(fhs[0][0])
        else :
            fh=0

        if not (szg==0 and fh==0) and rq!='-':
            rowdat = [dm,rq,szg,fh]
            data.append(rowdat)
        else:
            f.write("请检查第"+str(j)+"行，股票代码："+dm+"、日期："+rq+"、送转股："+str(szg)+"、分红："+str(fh)+"\r\n")
#            print("请检查第%d行，股票代码：%s、日期：%s、送转股：%f、分红：%f" % (j,dm,rq,szg,fh))

    dbcn.executemany('INSERT OR REPLACE INTO FH (GPDM,RQ,SZG,FH) VALUES (?,?,?,?)', data)
    dbcn.commit()
    dbcn.close()
    browser.quit()               
    f.close()

    return True

#########################################################################################
#提取指定报告期的分红数据
#########################################################################################   
def getdata1():
    bgq='201706'
    pg=3
    j=1
    while j<pg+1:
        if getdata(bgq,j,pg):
            j+=1
       
 
    
#########################################################################################
#提取数据
#########################################################################################   
def getdata2():
    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)

    bgqlist=[]
    for i in range(2017,1989,-1):
        for j in ('12','06'):
            bgq=str(i)+j
            if bgq in ('199006','201712'):
                continue
            else:
                bgqlist.append(bgq)
                
    bgq='199906'
    bg=bgqlist.index(bgq)
                
    for i in range(bg,len(bgqlist)):
    
        bgq=bgqlist[i]
        pg=getpgs(bgq)          
        print('报告期：%s共有%d页。' %(bgq,pg))

        j=1
        while j<pg+1:
            if getdata(bgq,j,pg):
                j+=1

    now2 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    print('结束运行时间：%s' % now2)

if __name__ == "__main__":  
#    createDataBase()
#    getdata2()    
     getdata1()  
    
        

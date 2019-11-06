# -*- coding: utf-8 -*-
"""
功能：本程序从股吧下载研报
用法：每天运行
"""
import datetime
import time
from selenium import webdriver
import sqlite3
import sys
import os
import re
from urllib import request


###############################################################################
#长股票代码
###############################################################################
def lgpdm(dm):
    return dm[:6]+('.SH' if dm[0]=='6' else '.SZ')

###############################################################################
#短股票代码
###############################################################################
def sgpdm(dm):
    return dm[:6]

########################################################################
#建立数据库
########################################################################
def createDataBase():
    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
    cn = sqlite3.connect(dbfn)

    '''
    gddm股东代码,gdmc股东名称,gdlx股东类型,gdpm股东排名,gpdm股票代码,gpmc股票名称,
    bgq报告期,cgsl持股数量,ltzb持股占流通股比例,zjsl持股增减数量,cgbd持股变动,ggrq公告日期
    '''
    cn.execute('''CREATE TABLE IF NOT EXISTS GDFX
           (GDDM TEXT NOT NULL,
           GDMC TEXT,
           GDLX TEXT,
           GDPM TEXT,
           GPDM TEXT NOT NULL,
           GPMC TEXT,
           BGQ TEXT NOT NULL,
           CGSL REAL,
           LTZB REAL,
           ZJSL REAL,
           CGBD TEXT,
           GGRQ TEXT NOT NULL);''')
    
    cn.execute('''CREATE UNIQUE INDEX IF NOT EXISTS GDFX_GPDM_GPDM_BGQ ON GDFX(GDDM,GPDM,BGQ);''')


def getdrive():
    return sys.argv[0][:2]

###############################################################################
#
###############################################################################

'''
 urllib.urlretrieve 的回调函数：
def callbackfunc(blocknum, blocksize, totalsize):
    @blocknum:  已经下载的数据块
    @blocksize: 数据块的大小
    @totalsize: 远程文件的大小
'''
 
def Schedule(blocknum, blocksize, totalsize):

    n = 10
    blk = int((totalsize / blocksize + (n-1))/n)
    blkn=[]
    for i in range(n):
        blkn.append(i*blk)
    
    if blocknum in blkn:   
        if blocknum==0:
            print('\n')
            
        recv_size = blocknum * blocksize
        speed = recv_size / (time.time() - start_time)
        # speed_str = " Speed: %.2f" % speed
        speed_str = " Speed: %s" % format_size(speed)
         
        # 设置下载进度条
    
        pervent = recv_size / totalsize
        
        percent_str = "%.2f%%" % (pervent * 100)
        
        
        n = round(pervent * 50)
        s = ('#' * n).ljust(50, '-')
        print(percent_str.ljust(8, ' ') + '[' + s + ']' + speed_str) 
        
    if blocknum >= totalsize/blocksize:
        print("100.00% "+"["+"#"*50+"] OK")

########################################################################
# 字节bytes转化K\M\G
########################################################################
def format_size(bytes):
    try:
        bytes = float(bytes)
        kb = bytes / 1024
    except:
        print("传入的字节格式不对")
        return "Error"
    if kb >= 1024:
        M = kb / 1024
        if M >= 1024:
            G = M / 1024
            return "%.3fG" % (G)
        else:
            return "%.3fM" % (M)
    else:
        return "%.3fK" % (kb)


if __name__ == "__main__": 
    print('%s Running' % sys.argv[0])

    start_time = time.time()

    data=[]
    gpdm='002675'
    dldir=r'D:\公司研究\东诚药业'
    browser = webdriver.PhantomJS()

    url0='http://guba.eastmoney.com'
    url='http://guba.eastmoney.com/list,%s,3,f.html' % gpdm
    browser.get(url)
    time.sleep(3)
    
    try:
        pgs=int(browser.find_element_by_class_name("sumpage").text)
    except:
        pgs=1
    
    if pgs>1 :

        for j in range(1,pgs+1):

            url='http://guba.eastmoney.com/list,%s,3,f_%d.html' % (gpdm,j)
            browser.get(url)
            time.sleep(3)
        
            elem = browser.find_element_by_id('articlelistnew')
            
        #    html=elem.get_attribute("outerHTML")
        
            rows=elem.find_elements_by_class_name("articleh")
            
            for i in range(len(rows)):
                
                td=rows[i].find_element_by_class_name("l3") 
                ybtitle=td.text
                
                yburl=td.find_element_by_tag_name('a').get_attribute("href")
            
                ybdate=rows[i].find_element_by_class_name("l6").text 
        
                data.append([ybtitle,yburl,ybdate])    
                
    else:

        elem = browser.find_element_by_id('articlelistnew')
            
        rows=elem.find_elements_by_class_name("articleh")
        
        for i in range(len(rows)):
            
            td=rows[i].find_element_by_class_name("l3") 
            ybtitle=td.text
            
            yburl=td.find_element_by_tag_name('a').get_attribute("href")
        
            ybdate=rows[i].find_element_by_class_name("l6").text 
    
            data.append([ybtitle,yburl,ybdate])    
           
    for i in range(742,len(data)):
        ybtitle = data[i][0]
        yburl = data[i][1]
        ybdate  = data[i][2]

        now1 = datetime.datetime.now().strftime('%H:%M:%S')

        print('当前时间：%s，正在处理第%d条: %s' % (now1,i+1,ybtitle))


        pdf_file=None

        txt_file=None
        
        browser = webdriver.PhantomJS()
        try:    
            
            browser.get(yburl)
            time.sleep(3)
            
            try:
                ybdate=browser.find_element_by_class_name("publishdate").text
    
                elem=browser.find_element_by_class_name("zwtitlepdf")
                pdfurl=elem.find_element_by_tag_name('a').get_attribute("href")
    
                pdfexist=True
            except:
                pdfexist=False
        
            if pdfexist:
                dlfn='['+ybdate+'] '+re.sub('[/:,*?"<>|]','_',ybtitle)+'.pdf'
                
                pdf_file= os.path.join(dldir,dlfn)
        
                if not os.path.exists(pdf_file):
                    print("正在下载公告 -- %s" % pdf_file)
                    '''
                    特别提醒：下调语句参数Schedule很重要，省略可能会出现出现被挂起，无响应的情况
                    '''
                    request.urlretrieve(pdfurl, pdf_file, Schedule)
            else:
                dlfn='['+ybdate+'] '+re.sub('[/:,*?"<>|]','_',ybtitle)+'.txt'
                
                txt_file= os.path.join(dldir,dlfn)
                txt=browser.find_element_by_class_name("stockcodec").text
                if not os.path.exists(txt_file):
                    print("正在下载公告 -- %s" % txt_file)
                    with open(txt_file,"w") as f:
                        f.write(txt)                            

        except:
            pass
        
        browser.quit()
        
            

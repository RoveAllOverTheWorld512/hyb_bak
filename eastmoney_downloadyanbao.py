# -*- coding: utf-8 -*-
"""
功能：本程序从东方财富网下载研报
用法：每天运行

金融界研报
http://istock.jrj.com.cn/yanbao_600674.html

"""
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
    gpdm='600674'
    dldir=r'D:\公司研究\川投能源'
    browser = webdriver.PhantomJS()

    url0='http://guba.eastmoney.com'
    url = 'http://data.eastmoney.com/report/%s.html' % gpdm

    browser.get(url)
    time.sleep(3)
    
    try:
        elemlst=browser.find_element_by_class_name("Page").text.split('\n')
        pgs=int(elemlst[len(elemlst)-4])
        
#        elem=browser.find_element_by_css_selector('div#PageCont a[title="转到最后一页"]')
#        html=elem.get_attribute("outerHTML")
#
#        pgshtml=browser.find_element_by_link_text("下一页")        

    except:
        pgs=1
    
    if pgs>1 :

        for j in range(1,pgs+1):

            browser.get(url)
            time.sleep(3)
        
            elem = browser.find_element_by_id("gopage")
            elem.clear()
            #输入页面
            elem.send_keys(j)
            elem = browser.find_element_by_class_name("btn_link")     
            #点击Go
            elem.click()
            time.sleep(5)
            
            elem= browser.find_element_by_id("dt_1")
            
            rows=elem.find_elements_by_tag_name("ul")
            
            for i in range(len(rows)):
                
                tds=rows[i].find_elements_by_tag_name("li") 
                
                ybtitle=tds[4].text
                
                yburl=tds[4].find_element_by_tag_name('a').get_attribute("href")
            
                ybdate='研报日期：'+tds[0].text 
        
                data.append([ybtitle,yburl,ybdate])    
                
    else:

            elem= browser.find_element_by_id("dt_1")
            
            rows=elem.find_elements_by_tag_name("ul")
            
            for i in range(len(rows)):
                
                tds=rows[i].find_elements_by_tag_name("li") 
                
                ybtitle=tds[4].text
                
                yburl=tds[4].find_element_by_tag_name('a').get_attribute("href")
            
                ybdate='研报日期：'+tds[0].text 
        
                data.append([ybtitle,yburl,ybdate])    
           
    for i in range(0,len(data)):

        ybtitle = data[i][0]
        yburl = data[i][1]
        ybdate  = data[i][2]

        print(ybtitle)

        pdf_file=None

        txt_file=None
        
#        browser = webdriver.PhantomJS()
        try:    
            
            browser.get(yburl)
            time.sleep(3)
            
            try:
    
                elem=browser.find_element_by_class_name("report-infos")
                
                span4=elem.find_elements_by_tag_name('span')[4]
                
                pdfurl=span4.find_element_by_tag_name('a').get_attribute("href")
    
                pdfexist=True
            except:
                pdfexist=False
        
            if pdfexist:
                dlfn='['+ybdate+'] '+re.sub('[/:,*?"<>|]','_',ybtitle)+'.pdf'
                
                pdf_file= os.path.join(dldir,dlfn)
        
                if not os.path.exists(pdf_file):
                    print("正在下载研报 -- %s" % pdf_file)
                    '''
                    特别提醒：下调语句参数Schedule很重要，省略可能会出现出现被挂起，无响应的情况
                    '''
                    request.urlretrieve(pdfurl, pdf_file, Schedule)
            else:
                dlfn='['+ybdate+'] '+re.sub('[/:,*?"<>|]','_',ybtitle)+'.html'
                
                txt_file= os.path.join(dldir,dlfn)

                html1='<html><head><style type="text/css">.newsContent {text-indent:2em ;}body {font-size: 24px;font-weight:normal;font-family:"宋体";}</style></head><body>'

                html=browser.find_element_by_class_name("newsContent").get_attribute("outerHTML")
                html2='</body></html>'
                
                if not os.path.exists(txt_file):
                    print("正在下载研报 -- %s" % txt_file)
                    with open(txt_file,"w") as f:
                        f.write(html1+html+html2)                            

        except:
            pass
        
    browser.quit()
        
            

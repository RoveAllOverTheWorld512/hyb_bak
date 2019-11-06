# -*- coding: utf-8 -*-
"""
本程序从东方财富网提取股东户数的最新变化情况，保存文件名为提取“股东户数+提取日期”
"""
import time
import datetime
from selenium import webdriver
import os
import sys
import struct
import pandas as pd
import xlwings as xw

def getdisk():
    return sys.argv[0][:2]

def makedir(dirname):
    if dirname == None :
        return False

    if not os.path.exists(dirname):
        try :
            os.mkdir(dirname)
            return True
        except(OSError):
            print("创建目录%s出错，请检查！" % dirname)
            return False
    else :
        return True

def filename(pathname):
    wjm = os.path.splitext(os.path.basename(pathname))
    return wjm[0]


#############################################################################
#读取dbf文件
#############################################################################
def dbfreader(f):

    numrec, lenheader = struct.unpack('<xxxxLH22x', f.read(32))
    numfields = (lenheader - 33) // 32

    fields = []
    for fieldno in range(numfields):
        name, typ, size, deci = struct.unpack('<11sc4xBB14x', f.read(32))
        name = name.decode().replace('\x00', '')
        typ  = typ.decode()
        fields.append((name, typ, size, deci))
    yield [field[0] for field in fields]
    yield [tuple(field[1:]) for field in fields]

    terminator = f.read(1)

    fields.insert(0, ('DeletionFlag', 'C', 1, 0))
    fmt = ''.join(['%ds' % fieldinfo[2] for fieldinfo in fields])
    fmtsiz = struct.calcsize(fmt)

    for i in range(numrec):
        record = struct.unpack(fmt, f.read(fmtsiz))
        if record[0].decode() != ' ':
            continue                        # deleted record
        result = []
        for (name, typ, size, deci), value in list(zip(fields, record)):
            if name == 'DeletionFlag':
                continue
            if typ == "C":
                value = value.strip(b'\x00').decode('GBK')
            if typ == "N":
                value = value.strip(b'\x00').strip(b'\x20').decode('GBK')
                if value == '':
                    value = 0
                elif deci:
                    value = float(value)
                else:
                    value = int(value)
            elif typ == 'D':
                value = value.decode('GBK')
            elif typ == 'L':
                value = value.decode('GBK')
                value = (value in 'YyTt' and 'T') or (value in 'NnFf' and 'F')

            result.append(value)

        yield result

def dbf2pandas(dbffn,cols):
    with open(dbffn,"rb") as f:
        data = list(dbfreader(f))
        f.close()
    columns = data[0]
    columns=[e.lower() for e in columns]
    data = data[2:]
    df = pd.DataFrame(data,columns=columns)
    if len(cols) == 0 :
        return df
    else :
        return df[cols]

##########################################################################
#将字符串转换为时间戳，不成功返回None
##########################################################################
def str2datetime(s):
    try:
        dt = datetime.datetime(int(s[:4]),int(s[4:6]),int(s[6:8]))
    except :
        dt = None
    return dt

##########################################################################
#将日期"20170101"字符串转换成“2017-01-01”，先检测是不是日期不，成功返回None
##########################################################################
def datefmt(s):
    try:
        dt = datetime.datetime(int(s[:4]),int(s[4:6]),int(s[6:8]))
        dt = s[:4]+'-'+s[4:6]+'-'+s[6:8]
    except :
        dt = None
    return dt

##########################################################################
#提取上市公司的上市时间
##########################################################################
def ssdate():
    fn=r"C:\new_hxzq_hc\T0002\hq_cache\base.dbf"
    sssj = dbf2pandas(fn,['gpdm', 'ssdate', 'sc']) 
    sssj['ssdate'] = sssj['ssdate'].map(datefmt)
    return sssj

########################################################################
#检测是不是可以转换成浮点数
########################################################################
def str2float(num):
    try:
        return float(num)
    except ValueError:
        return num

########################################################################
#生成Excel文件
########################################################################
if __name__ == "__main__": 
    
    now1 = datetime.datetime.now().strftime('%H:%M:%S')

    #提取上市公司的上市时间    
    fn=r"C:\new_hxzq_hc\T0002\hq_cache\base.dbf"
    sssj = dbf2pandas(fn,['gpdm', 'ssdate']) 
    sssj['ssdate'] = sssj['ssdate'].map(datefmt)
    
    today = datetime.datetime.now().strftime("%Y%m%d")    

    fld1=['股票代码','股票名称','股价(元)','本次户数','上次户数','增减户数','增加比例(%)',
                 '区间涨幅(%)','本次截止日期','上次截止日期','户均市值(万)','户均股数(万)',
                 '总市值(亿)','总股本(亿)','公告日期','上市日期']
    fld2=['gpdm','gpmc','gj','bchs','schs','zjhs','zjbl',
                 'qjzf','bcjzrq','scjzrq','hjsz','hjgs',
                 'zsz','zgb','ggrq']

#    browser = webdriver.Firefox()
    browser = webdriver.PhantomJS()
    browser.maximize_window()
    
    browser.get("http://data.eastmoney.com/gdhs/")
    
    data = []
    for k in range(0,6):
        print("正在处理第%d页，请等待。" % k)
        elem = browser.find_element_by_id("PageContgopage")
        elem.clear()
        elem.send_keys(k)
        elem = browser.find_element_by_class_name("btn_link")        
        elem.click()
        time.sleep(2)
        tbody = browser.find_elements_by_tag_name("tbody")
        tblrows = tbody[0].find_elements_by_tag_name('tr')
       
        for j in range(len(tblrows)):
            rowdat = []
            tblcols = tblrows[j].find_elements_by_tag_name('td')
            for i in range(len(tblcols)):
                if i not in (2,4):
                    if i in (11,16):
                        coldat = tblcols[i].find_elements_by_tag_name('span')[0].get_property("title")
                    else:
                        coldat = tblcols[i].text
                    if i not in (0,1,10,11,16):
                        rowdat.append(str2float(coldat))
                    else:
                        rowdat.append(coldat)
            data.append(rowdat)
    
    browser.quit()
    
    gdhs = pd.DataFrame(data,columns=fld2)
    gdhs1 = pd.merge(gdhs,sssj,on="gpdm")
    #剔除上市时间晚于2016年5月1日的
#    gdhs1=gdhs1[gdhs1['ssdate']<'2016-05-01']
    #剔除股东户数增加的
#    gdhs1=gdhs1[gdhs1['zjbl']<-3]
    #剔除三季报以前资料
#    gdhs1=gdhs1[gdhs1['bcjzrq']>'2017/09/30']
#    gdhs1=gdhs1[gdhs1['scjzrq']>='2017/09/30']
    #按增加比例升序排列
#    gdhs1=gdhs1.sort_values(by="zjbl")
    #去掉重复股票
    gdhs1.drop_duplicates(['gpdm','bcjzrq','scjzrq'],keep='last',inplace=True)
    
    gdhs1.columns = fld1
    
    fn = r'd:\\hyb\\股东户数_'+today+'.xlsx'   
    writer = pd.ExcelWriter(fn, engine='xlsxwriter')

    gdhs1.to_excel(writer, sheet_name='最新股东户数',index=False)

    workbook = writer.book
    worksheet = writer.sheets['最新股东户数']

#    format1 = workbook.add_format({'num_format': '0.0000'})
#    format2 = workbook.add_format({'num_format': '0.00'})
    format3 = workbook.add_format({'num_format': 'yyyy-mm-dd'})
    
    worksheet.set_column('A:P', 10)
    worksheet.set_column('I:J', 12, format3)
    worksheet.set_column('O:P', 12, format3)
    worksheet.freeze_panes(1, 0)

    writer.save()
    print('请打开%s文件查看最新股东户数股票名单。' % fn)

    now2 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    print('结束运行时间：%s' % now2)
#    return fn

########################################################################
#主程序
########################################################################
#if __name__ == "__main__": 
#    
#
#    today = datetime.datetime.now().strftime("%Y%m%d")    
#    xlfn = r'd:\\hyb\\股东户数_'+today+'.xlsx'   
##    if not os.path.exists(xlfn):
#    xlfn=createexcel()
    
#    app=xw.App(visible=True,add_book=False)
#    app.display_alerts=False
#    app.screen_updating=False
#
#    wb=app.books.open(xlfn)
#    sh=wb.sheets['最新股东户数']
#    rng=xw.Range('A1')
#    rows=rng.current_region.shape[0]
#    for i in range(2,rows+1):
#        add='A'+str(i)
#        rng=xw.Range(add)
#        vl=rng.value
#        vl1=vl[0:6]+('.SH' if vl[0]=='6' else '.SZ')
#        rng.number_format='@'
#        lnk=r'http://data.eastmoney.com/gdhs/detail/'+vl[0:6]+r'.html'
#        rng.add_hyperlink(lnk,vl1,'提示：点击即链接到东方财富网该股票股东户数历史详情')
#        add='B'+str(i)
#        rng=xw.Range(add)
#        vl1=rng.value
#        lnk=r'http://data.eastmoney.com/report/'+vl[0:6]+r'.html'
#        rng.add_hyperlink(lnk,vl1,'提示：点击即链接到东方财富网该股票研报')
#        

    
#    wb.save()
        

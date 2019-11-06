# -*- coding: utf-8 -*-
"""
本程序将cninfo网站的分红配股信息抓取保存到“d:\公司研究\个股名称”目录下
分红配股[股票代码][股票名称].dbf

http://www.cninfo.com.cn/information/allotment/szcn300121.html
http://www.cninfo.com.cn/information/allotment/szmb000001.html
http://www.cninfo.com.cn/information/allotment/shmb600000.html

"""
import re
import os
import struct
import unicodedata
import datetime
from pyquery import PyQuery as pq
import winreg
import sys

#############################################################################
#返回包含中文的byte字符串转的长度(一个汉字的长度为2)
#############################################################################
def str_width(s):
    w=0
    for c in s:
        if (unicodedata.east_asian_width(c) in ('F','W')):
            w +=2
        else:
            w +=1
    return(w)

#############################################################################
#将包含中文的byte字符串转变为指定长度（一个汉字为2个宽度,后面用空格补齐)
#############################################################################
def cnstrjust(cnstr,length):
    cnstrw=str_width(cnstr)
    if cnstrw>length :
        i=0
        while i<len(cnstr):
            i += 1
            cutstr = cnstr[:i]
            if str_width(cutstr)>length :
                break
        cnstr = cutstr[:i-1]
        cnstrw=str_width(cnstr)

    return cnstr+" "*(length-cnstrw)

#############################################################################
#检查字段类型和宽度，如果N，C型宽度不够，则扩展宽度，如果D，L型宽度不符则改为C型
#############################################################################
def checkdata(fieldnames, fieldspecs, records):
    flds = []
    for name, (typ, size, deci) in list(zip(fieldnames, fieldspecs)):
        i = fieldnames.index(name)
        if typ=="N" :
            p="{:>"+str(size)+"."+str(deci)+"f}"
            maxlen = max([len(p.format(e[i])) for e in records])
            if maxlen>size :
                size = maxlen
        if typ=="C" :
            maxlen = max([len(e[i]) for e in records])
            if maxlen>size :
                size = maxlen
        if typ in ("D","L"):
            maxlen = max([len(e[i]) for e in records])
            minlen = min([len(e[i]) for e in records])
            if maxlen!=8 or minlen!=8 :
                typ = 'C'
                size = maxlen
        flds.append([typ,size,deci])
    return flds

#############################################################################
#写整个dbf文件
#############################################################################
def dbfwriter(f, fieldnames, fieldspecs, records):
    #对数据与字段的类型和宽度进行检查、优化
    if len(records)>0 :
        fieldspecs = checkdata(fieldnames, fieldspecs, records)
    # 文件头部信息
    ver = 3
    now = datetime.datetime.now()
    yr, mon, day = now.year-2000, now.month, now.day
    numrec = len(records)
    numfields = len(fieldspecs)
    lenheader = numfields * 32 + 33
    lenrecord = sum(field[1] for field in fieldspecs) + 1
    codepageid = 122
    #Code Pages Supported by Visual FoxPro:936Chinese (PRC, Singapore) Windows
    #https://technet.microsoft.com/zh-cn/learning/aa975345
    hdr = struct.pack('<BBBBLHH17xB2x', ver, yr, mon, day, numrec, lenheader, lenrecord, codepageid)
    f.write(hdr)

    # 字段名信息
    addr = 1
    for name, (typ, size, deci) in list(zip(fieldnames, fieldspecs)):
        name = name.ljust(11, '\x00').encode('GBK')
        typ = typ.encode('GBK')
        fld = struct.pack('<11sciBB14x', name, typ, addr, size, deci)
        addr += size
        f.write(fld)

    # 终止符
    f.write('\r'.encode())

    # 记录
    for record in records:
        f.write(' '.encode())                        # deletion flag
        for (typ, size, deci), value in list(zip(fieldspecs, record)):
            if typ == "C":
                value = cnstrjust(value,size)
            if typ == "N":
                p="{:>"+str(size)+"."+str(deci)+"f}"
                value = p.format(value)

            if typ == 'D':
                value = value.ljust(8, ' ')
            if typ == 'L':
                value = value.upper()

            f.write(value.encode("GBK"))

    # 文件尾
    f.write('\x1A'.encode())

######################################################################################
#生成分红配股DBF文件
######################################################################################    
"""
gpdm股票代码
gpmc股票名称
gqdjr股权登记日
cqjzr除权登记日
hgssr红股上市日
mgfh每股分红
mgsg每股送股
mgpg每股配股
pgj配股价
bz备注，用来保存抓取的信息
bj标记,标记为1表明有分红、也有送股、配股等多个方案，标记为2表明分红配股信息须人工识别
"""
def dbfcreate(dbffn,records):
#    dbffn = 'fhpg.dbf'

    fieldnames = ['gpdm','gpmc','gqdjr', 'cqjzr', 'hgssr','mgfh',
                  'mgsg','mgpg','pgj','bz','bj']
    fieldspecs = [('C', 6, 0),('C', 8, 0),('D', 8, 0),('D', 8, 0),('D', 8, 0),('N', 8, 5),
                  ('N', 8, 5),('N', 8, 5),('N', 8, 5),('C', 90, 0),('C', 1, 0)]
#    records = []

    with open(dbffn,"wb") as f:
        dbfwriter(f, fieldnames, fieldspecs, records)
        f.close()

######################################################################################
#检测路径是否存在，不存则创建
######################################################################################    
def exsit_path(pth):
    if not os.path.exists(pth) :
        os.makedirs(pth)

######################################################################################
#从cninfo提取分红配股信息
######################################################################################    
"""

"""
def tqfhpg(gpdm):
    urlpref = "http://www.cninfo.com.cn/information/dividend/"
    if gpdm[0] == '6' :
        url = urlpref + "shmb"+gpdm+".html"
    elif gpdm[:3] == "002" :
        url = urlpref + "szsme"+gpdm+".html"
    elif gpdm[:3] == "300" :
        url = urlpref + "szcn"+gpdm+".html"
    elif gpdm[:3] in ("000","001") :
        url = urlpref + "szmb"+gpdm+".html"
    else :
        url = ""

    html = pq(url,encoding="gb2312")
    tb = html('tr')
    gpmc=pq(html('tr').eq(0).html()).find('td').eq(0).text().strip()[-4:]
    
    gpmc=gpmc.replace(" ","").replace("*","")

    pth = 'd:\\公司研究\\'+gpmc
    
    exsit_path(pth)
    fn = pth+'\\'+gpdm+gpmc+'分红配股.dbf'
    if os.path.exists(fn):
        os.remove(fn)
    
    
    records = []
    if len(tb)<3 :
        return 

    for i in range(3,len(tb)) :

        fh = 0     #每个分红
        sg = 0     #每股送股和转增股数
        pg = 0     #每股配股
        pgj = 0    #配股价
        fas = 0    #方案数
        bj = "1"
        row=pq(html('tr').eq(i).html())
        fhfa = row.find('td').eq(1).text()    #分红方案
        fhfa = fhfa.encode('gbk','ignore').decode('gbk','ignore')
        fhstr = row.find('td').eq(1).text()    #分红方案
        fhstr = fhstr.encode('gbk','ignore').decode('gbk','ignore')
        gqdjr = row.find('td').eq(2).text()    #股权登记日
        gqdjr = gqdjr.strip()
        cqjzr = row.find('td').eq(3).text()    #除权基准日
        cqjzr = cqjzr.strip()
        hgssr = row.find('td').eq(3).text()    #红股上市日
        hgssr = hgssr.strip()

        fhstr = fhstr.replace(" ","")
        fhstr = fhstr.replace("股","")
        fhstr = fhstr.replace("赠","增")
        fhstr = fhstr.replace("元","")

        if len(fhstr) == 0:
            continue

        fhs = re.findall ('([\d\.]+)派([\d\.]+)',fhstr)
        if len(fhs) >1 :
            print(fhstr)
        elif len(fhs) == 1:
            fh = float(fhs[0][1])/float(fhs[0][0])
            sg = 0
            fas += 1
            bj = ""

        fhs = re.findall ('([\d\.]+)送([\d\.]+)',fhstr )
        if len(fhs) >1 :
            print(fhstr)
        elif len(fhs) == 1:
            fh = 0
            sg = float(fhs[0][1])/float(fhs[0][0])
            fas += 1
            bj = ""

        fhs = re.findall ('([\d\.]+)转增([\d\.]+)',fhstr )
        if len(fhs) >1 :
            print(fhstr)
        elif len(fhs) == 1:
            fh = 0
            sg = float(fhs[0][1])/float(fhs[0][0])
            fas += 1
            bj = ""

        fhs = re.findall ('([\d\.]+)转([\d\.]+)',fhstr )
        if len(fhs) >1 :
            print(fhstr)
        elif len(fhs) == 1:
            fh = 0
            sg = float(fhs[0][1])/float(fhs[0][0])
            fas += 1
            bj = ""

        fhs = re.findall ('([\d\.]+)送([\d\.]+)派([\d\.]+)',fhstr )
        if len(fhs) >1 :
            print(fhstr)
        elif len(fhs) == 1:
            fh = float(fhs[0][2])/float(fhs[0][0])
            sg = float(fhs[0][1])/float(fhs[0][0])
            fas += 1
            bj = ""

        fhs = re.findall ('([\d\.]+)转增([\d\.]+)派([\d\.]+)',fhstr )
        if len(fhs) >1 :
            print(fhstr)
        elif len(fhs) == 1:
            fh = float(fhs[0][2])/float(fhs[0][0])
            sg = float(fhs[0][1])/float(fhs[0][0])
            fas += 1
            bj = ""

        fhs = re.findall ('([\d\.]+)送([\d\.]+)转增([\d\.]+)',fhstr )
        if len(fhs) >1 :
            print(fhstr)
        elif len(fhs) == 1:
            fh = 0
            sg = (float(fhs[0][1])+float(fhs[0][2]))/float(fhs[0][0])
            fas += 1
            bj = ""

        fhs = re.findall ('([\d\.]+)转增([\d\.]+)送([\d\.]+)',fhstr )
        if len(fhs) >1 :
            print(fhstr)
        elif len(fhs) == 1:
            fh = 0
            sg = (float(fhs[0][1])+float(fhs[0][2]))/float(fhs[0][0])
            fas += 1
            bj = ""

        fhs = re.findall ('([\d\.]+)送([\d\.]+)转增([\d\.]+)派([\d\.]+)',fhstr )
        if len(fhs) >1 :
            print(fhstr)
        elif len(fhs) == 1:
            fh = float(fhs[0][3])/float(fhs[0][0])
            sg = (float(fhs[0][1])+float(fhs[0][2]))/float(fhs[0][0])
            fas += 1
            bj = ""

        fhs = re.findall ('([\d\.]+)转增([\d\.]+)送([\d\.]+)派([\d\.]+)',fhstr )
        if len(fhs) >1 :
            print(fhstr)
        elif len(fhs) == 1:
            fh = float(fhs[0][3])/float(fhs[0][0])
            sg = (float(fhs[0][1])+float(fhs[0][2]))/float(fhs[0][0])
            fas += 1
            bj = ""

        if fas > 1 :
            bj = "1"
            
        if (fh==0 and sg==0 and sg==0) :
            bj = "2"

        records.append([gpdm,gpmc,gqdjr,cqjzr,hgssr,fh,sg,pg,pgj,fhfa,bj])
        
    dbfcreate(fn,records)

    return


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


if __name__ == '__main__':
    zxgfile="zxg.blk"
    tdxblkdir = gettdxblkdir()
    zxgfile = os.path.join(tdxblkdir,zxgfile)
    zxglb = zxglist(zxgfile,"tdxblk")
    j=154       #最小值为1
    k=154
    l=k if k<=len(zxglb) else len(zxglb)
    for i in range(j-1,l):
        gpdm=zxglb[i]
        print('共有%d只股票,正在处理第%d只股票，请等待。' %(len(zxglb),i+1))
        tqfhpg(gpdm)
    
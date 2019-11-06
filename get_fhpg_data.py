# -*- coding: utf-8 -*-
"""
Created on Thu Feb  11 14:40:41 2017

@author: huangyunbin@sina.com
"""

import re
import os
import struct
import unicodedata
import datetime
from pyquery import PyQuery as pq
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
                value = value.strip(b'\x00').decode('GBK')
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

def dbfcreate():
    dbffn = 'fhpg.dbf'

    fieldnames = ['gpdm','gqdjr', 'cqjzr', 'hgssr','mgfh',
                  'mgsg','mgpg','pgj','bz','bj']
    fieldspecs = [('C', 6, 0),('D', 8, 0),('D', 8, 0),('D', 8, 0),('N', 8, 5),
                  ('N', 8, 5),('N', 8, 5),('N', 8, 5),('C', 90, 0),('C', 1, 0)]
    records = []

    with open(dbffn,"wb") as f:
        dbfwriter(f, fieldnames, fieldspecs, records)
        f.close()


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
    fhpgda = []
    if len(tb)<3 :
        return fhpgda

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

        fhpgda.append([gpdm,gqdjr,cqjzr,hgssr,fh,sg,pg,pgj,fhfa,bj])

    return fhpgda

def dbffhpg(records):

    fhpgda = []
    for [gpdm,gqdjr,cqjzr,hgssr,mgfh,mgsg,mgpg,pgj,fhfa,bj] in records :

        fh = 0     #每个分红
        sg = 0     #每股送股和转增股数
        pg = 0     #每股配股
        pgj = 0    #配股价
        fas = 0    #方案数
        bj = "1"

        fhstr = fhfa
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

        fhpgda.append([gpdm,gqdjr,cqjzr,hgssr,fh,sg,pg,pgj,fhfa,bj])

    return fhpgda

def getfhpg(dm):

    with open('gpdmb.txt') as f:
        dmb=f.read()
        f.close()
    gpdms = re.findall('(\d{6})',dmb)
    if len(dm)>0 :
        bg = gpdms.index(dm)+1
    else :
        bg =0

    fn = 'fhpg.dbf'
    appendrds = []
    for i in range(bg,len(gpdms)):

        gpdm = gpdms[i]

        print("共有%d个，正在处理第%d个：%s" % (len(gpdms),i+1,gpdm))
        appendrds=(tqfhpg(gpdm))
        if len(appendrds)==0 :
            continue

        with open(fn,'rb') as f:
            olddata = list(dbfreader(f))
            f.close()
        if os.path.exists(fn+".bak") :
            os.remove(fn+".bak")
        os.rename(fn,fn+".bak")

        fieldnames = olddata[0]
        fieldspecs = olddata[1]
        records = olddata[2:]
        records.extend(appendrds)
        with open(fn,'wb') as f:
            dbfwriter(f, fieldnames, fieldspecs, records)
            f.close()

###############################################################################
#读取提前的最后一个股票代码
###############################################################################
def getlastdm():
    fn = 'fhpg.dbf'
    with open(fn,'rb') as f:
        olddata = list(dbfreader(f))
        f.close()

    records = olddata[2:]
    if len(records) > 0 :
        return records[len(records)-1][0]
    else :
        return ''
    

def tmp():

    fn = 'fhpg.dbf'
    with open(fn,'rb') as f:
        olddata = list(dbfreader(f))
        f.close()

    fieldnames = olddata[0]
    fieldspecs = olddata[1]
    records = olddata[2:]

    records = dbffhpg(records)

    if os.path.exists(fn+".bak") :
        os.remove(fn+".bak")
    os.rename(fn,fn+".bak")

    with open(fn,'wb') as f:
        dbfwriter(f, fieldnames, fieldspecs, records)
        f.close()

if __name__ == '__main__':
    fn = 'fhpg.dbf'
    if os.path.exists(fn) : 
        dm = getlastdm()
    else :
        dbfcreate()
        
    getfhpg(dm)
#    tmp()

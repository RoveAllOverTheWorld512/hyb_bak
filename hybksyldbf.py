# -*- coding: utf-8 -*-
"""
Created on Fri Jan 20 11:27:04 2017

@author: lenovo
"""
import re
import os
import sys
import struct
import unicodedata
import datetime
import xlrd

########################################################################
#检测是不是可以转换成浮点数
########################################################################
def isfloat(num):
    try:
        float(num)
        return True
    except ValueError:
        return False


def readhysyl(file,sheet,rowbg):
    wb = xlrd.open_workbook(file,encoding_override="cp1252")
    table = wb.sheet_by_name(sheet)
    nrows = table.nrows #行数

    data =[]
    for rownum in range(rowbg,nrows):
        row = table.row_values(rownum)
        data.append([row[0],row[1],float(row[2]) if isfloat(row[2]) else 0])
    return data

def readbksyl(file,sheet,rowbg):
    wb = xlrd.open_workbook(file,encoding_override="cp1252")
    table = wb.sheet_by_name(sheet)
    nrows = table.nrows #行数
    dmzd = {"上海A股":"SHAG","深圳A股":"SZAG","沪深A股":"HSAG",
            "深市主板":"SSZB","中小板":"ZXB","创业板":"CYB"}
    data =[]
    for rownum in range(rowbg,nrows):
        row = table.row_values(rownum)
        data.append(['0',dmzd[row[0]],row[0],float(row[1]) if isfloat(row[1]) else 0])
    return data


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
#写dbf文件
#############################################################################
def dbfwriter(f, fieldnames, fieldspecs, records):
    #对数据与字段的类型和宽度进行检查、优化
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


##########################################################################
#获取运行程序所在驱动器
##########################################################################
def getdrive():
    return os.path.splitdrive(sys.argv[0])[0]


def hysylrd():
    syldir = getdrive()+r'\syl'
    files = os.listdir(syldir)
    fs = [re.findall('csi(\d{8})\.xls',e) for e in files]
    jyrqlist =[]
    for e in fs:
        if len(e)>0:
            jyrqlist.append(e[0])
    records = []
    for rq in jyrqlist:
        xlsfn=syldir + "\\csi"+rq+".xls"
        rds = readhysyl(xlsfn,'中证行业滚动市盈率',1)
        rd = [e.append(rq) for e in rds]
        rd = [e.insert(0,str(int(len(e[0])/2))) for e in rds]
        records.extend(rds)

    return records

def bksylrd():
    syldir = getdrive()+r'\bksyl'
    files = os.listdir(syldir)
    fs = [re.findall('bk(\d{8})\.xls',e) for e in files]
    jyrqlist =[]
    for e in fs:
        if len(e)>0:
            jyrqlist.append(e[0])
    records = []
    for rq in jyrqlist:
        xlsfn=syldir + "\\bk"+rq+".xls"
        rds = readbksyl(xlsfn,'板块滚动市盈率',1)
        rd = [e.append(rq) for e in rds]
        records.extend(rds)

    return records

if __name__ == '__main__':
    records = bksylrd()
    records.extend(hysylrd())
    fieldnames = ['dmjb','hydm', 'hymc', 'syl','rq']
    fieldspecs = [('C', 1, 0),('C', 8, 0), ('C', 30, 0), ('N', 8, 2), ('C', 8, 0)]
    fn = 'hybksyl.dbf'

    f = open(fn,"wb")

    dbfwriter(f, fieldnames, fieldspecs, records)
    f.close()







# -*- coding: utf-8 -*-
"""
Created on Fri Feb  3 11:55:03 2017
功能:休市安排处理
@author: lenovo
"""
import datetime
import xlwings as xw
from configobj import ConfigObj  
import pandas as pd


##############################################################################
# 不可用type(123.456)=="float"判断
##############################################################################
def int2str(n):
    if isinstance(n,float):
        return str(int(n))
    if isinstance(n,int):
        return str(n)
    return n
        
##########################################################################
#将字符串转换为时间戳，不成功返回None
##########################################################################
def str2datetime(s):
    dt = None

    if ('-' in s) or ('/' in s):
        if '-' in s:
            dt=s.split('-')
        if '/' in s:
            dt=s.split('/')        
        try:
            dt = datetime.datetime(int(dt[0]),int(dt[1]),int(dt[2]))
        except :
            dt = None

    if len(s)==8:
        try:
            dt = datetime.datetime(int(s[:4]),int(s[4:6]),int(s[6:8]))
        except :
            dt = None

    return dt


if __name__ == '__main__':
    xlsfn=r'D:\hyb\2018年股市休市安排表.xlsx'

    wb = xw.Book(xlsfn)
    data = wb.sheets[0].range('B2:C2').options(expand='down').value
    clsdt=[]
    for row in data:
        d1=str2datetime(int2str(row[0]))
        d2=str2datetime(int2str(row[1]))
        rng=pd.bdate_range(d1,d2)
        clsdt.extend(rng)
    clsdt=[e.strftime("%Y%m%d") for e in clsdt]    
    conf_ini =r'd:\selestock\closedate.ini'
    config = ConfigObj(conf_ini,encoding='UTF8') 
    config['stockclosedate'] = str(clsdt)
    config.write()

     






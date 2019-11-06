# -*- coding: utf-8 -*-
"""
Created on Thu Nov 16 14:31:49 2017

@author: lenovo
"""
import pandas as pd

###############################################################################
#从通达信系统读取股票代码表
###############################################################################
def getcode():
    datacode = []
    for sc in ('h','z'):
        fn = r'C:\new_hxzq_hc\T0002\hq_cache\s'+sc+'m.tnf'
        f = open(fn,'rb')
        f.seek(50)
        ss = f.read(314)
        while len(ss)>0:
            gpdm=ss[0:6].decode('GBK')
            gpmc=ss[23:31].strip(b'\x00').decode('GBK')
            gppy=ss[285:289].strip(b'\x00').decode('GBK')
            #剔除非A股代码
            if gpdm[0] in ('6','3','0') :
                datacode.append([gpdm,gpmc,gppy])
            ss = f.read(314)
        f.close()
    gpdmb=pd.DataFrame(datacode,columns=['gpdm','gpmc','gppy'])
    return gpdmb

if __name__ == '__main__':
    gpdmb=getcode()
    
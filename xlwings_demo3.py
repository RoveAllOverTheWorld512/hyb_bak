# -*- coding: utf-8 -*-
"""
Created on Mon Nov 27 12:20:38 2017

@author: lenovo
"""

# 导入xlwings模块，打开Excel程序，默认设置：程序可见，只打开不新建工作薄，屏幕更新关闭
import xlwings as xw
app=xw.App(visible=True,add_book=False)
app.display_alerts=False
#app.screen_updating=False
# 文件位置：filepath，打开test文档，然后保存，关闭，结束程序
filepath=r'd:\hyb\股东户数_201711262.xlsx'
wb=app.books.open(filepath)

#app.screen_updating=True
sht1 = wb.sheets.add(name='gdhs')
sht2 = wb.sheets['最新股东户数']

rng2 = sht2.range('A1').expand('down') 
rng1 = sht1.range('A1')
data=rng2.options(ndim=2).value
for i in range(1,len(data)):
    dt=data[i][0]
    data[i][0]= dt + ('.SH' if dt[0]=='6' else '.SZ')
rng1.value=data


rng2 = sht2.range('I1').expand('down') 
rng1 = sht1.range('B1')
data=rng2.options(ndim=2).value
for i in range(1,len(data)):
    data[i][0]=data[i][0].replace('/','-')
rng1.value=data
rng1 = sht1.range('B1').expand('down') 
rng1.number_format="yyyy-mm-dd;@"

rng2 = sht2.range('D1').expand('down') 
rng1 = sht1.range('C1')
rng1.value=rng2.options(ndim=2).value

#wb.save()
#wb.close()
#app.quit()

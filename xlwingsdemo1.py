# -*- coding: utf-8 -*-
"""
Created on Fri Nov 17 11:04:49 2017

@author: lenovo
"""

 # 导入xlwings模块，打开Excel程序，默认设置：程序可见，只打开不新建工作薄，屏幕更新关闭
import xlwings as xw
app=xw.App(visible=True,add_book=False)
app.display_alerts=False
app.screen_updating=False
# 文件位置：filepath，打开test文档，然后保存，关闭，结束程序
filepath=r'd:\hyb\股东户数_20171117.xlsx'
wb=app.books.open(filepath)
wb.save()
wb.close()
app.quit()

app=xw.App(visible=True,add_book=False)
wb=app.books.add()
wb.save(r'd:\test.xlsx')
wb.close()
app.quit()
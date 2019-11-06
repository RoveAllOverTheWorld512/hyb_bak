# -*- coding: utf-8 -*-
"""
Created on Sun Dec 24 14:58:54 2017

@author: lenovo
"""

#coding:utf-8
import Tkinter
top=Tkinter.Tk()#创建顶层窗口
label=Tkinter.Label(top,text="hello \nworld")
label.pack()
quit=Tkinter.Button(top,text='quit',command=top.quit,bg='red',fg='white')
quit.pack(fill=Tkinter.X,expand=1)
Tkinter.mainloop()#加入服务
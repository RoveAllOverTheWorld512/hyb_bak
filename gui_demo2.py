# -*- coding: utf-8 -*-
"""
Created on Sun Dec 24 15:05:08 2017

@author: lenovo
"""

import tkinter as tk  
  
def button_callback(event):  
    print("ButtonPress")  
  
top = tk.Tk()   #创建主容器  
  
btn = tk.Button(top, text='click')       #创建按钮并设置父容器为top，设置按钮文字  
btn.pack(side='right')                 #将组件打包到父容器中显示，设置显示参数  
  
  
  
frame = tk.Frame(top)  #创建框架容器，可以在容器内布局  
  
sbar = tk.Scrollbar(frame).pack(side='left', fill='y')  
listbox = tk.Listbox(frame, height=15, width=50).pack(side='left', fill='both')  
  
frame.pack()  
  
  
btn.bind("<ButtonPress>", button_callback)  


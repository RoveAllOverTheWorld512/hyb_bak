# -*- coding: utf-8 -*-
"""
Created on Wed Feb 22 10:09:10 2017

@author: Lenovo
"""

#-*- coding: utf-8 -*-
from matplotlib.font_manager import FontProperties
import matplotlib.pyplot as plt
font = FontProperties(fname=r"c:\windows\fonts\simsun.ttc", size=14)
plt.figure(figsize=(6,6))


x = [1,2,3,4,5,6,7,8]
y = []
for i in x:
    y.append(-(i*i)+i+3)


plt.plot(x, y)
plt.title('测试程序', fontproperties=font)
plt.xlabel('x轴', fontproperties=font)
plt.ylabel('y轴', fontproperties=font)
plt.grid(True)
plt.show()
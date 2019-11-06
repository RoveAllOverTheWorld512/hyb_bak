# -*- coding: utf-8 -*-
"""
Created on Thu Feb 16 16:49:03 2017
Statsmodels 统计包之 OLS 回归
http://sanwen.net/a/ehleupo.html
matplotlib.pyplot入门
http://www.tuicool.com/articles/7zYNZfI
matplotlib绘图实例：pyplot、pylab模块及作图参数
http://blog.csdn.net/pipisorry/article/details/40005163
@author: Lenovo
"""

import numpy as np
import matplotlib.pyplot as plt

x = np.linspace(0, 10, 1000)
y = np.sin(x)
z = np.cos(x**2)

plt.figure(figsize=(8,4))
plt.plot(x,y,label="$sin(x)$",color="red",linewidth=2)
plt.plot(x,z,"b--",label="$cos(x^2)$")
plt.xlabel("Time(s)")
plt.ylabel("Volt")
plt.title("PyPlot First Example")
plt.ylim(-1.2,1.2)
plt.legend()
plt.show()
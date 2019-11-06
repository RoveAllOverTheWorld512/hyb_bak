# -*- coding: utf-8 -*-
"""
Created on Thu Mar  2 11:08:25 2017

@author: Lenovo
"""

import matplotlib.pyplot as plt
import numpy as np

t = np.linspace(0, 10, 1000)
y = np.sin(t)
plt.plot(t, y,label=u'正弦曲线 (m)')
plt.xlabel(u"时间", fontproperties='SimHei')
plt.ylabel(u"振幅", fontproperties='SimHei')
plt.title(u"正弦波", fontproperties='SimHei')

# 添加单位
t=plt.text(6.25, -1.14,r'$(\mu\mathrm{mol}$'+' '+'$ \mathrm{m}^{-2} \mathrm{s}^{-1})$',fontsize=15, horizontalalignment='center',verticalalignment='center')

#在这里设置是text的旋转，0为水平，90为竖直
t.set_rotation(0)

# legend中显示中文
plt.legend(prop={'family':'SimHei','size':15})

plt.savefig("test.png")
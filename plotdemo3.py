# -*- coding: utf-8 -*-
"""
Created on Wed Feb 22 16:29:38 2017

@author: Lenovo
"""

import matplotlib.pyplot as plt

fig = plt.figure()
ax = fig.add_subplot(111)
ax.plot([1,2,3,14],'ro-')

# set your ticks manually
ax.xaxis.set_ticks([1.,2.,3.,10.])
ax.grid(True)

plt.show()
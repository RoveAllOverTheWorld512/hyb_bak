# -*- coding: utf-8 -*-
"""
Created on Tue Feb 14 13:56:09 2017
http://blog.csdn.net/csqazwsxedc/article/details/51336322
@author: Lenovo
"""

import statsmodels.api as sm
import statsmodels.formula.api as smf
import statsmodels.graphics.api as smg
import patsy

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from pandas import Series,DataFrame
from scipy import stats
import seaborn as sns
import pandas_datareader as pdr

import datetime
start = datetime.datetime(2016,1,1)
end = datetime.datetime(2016,12,31)
datass = pdr.get_data_yahoo("000001.SS",start,end)
datajqr = pdr.get_data_yahoo("300024.SZ",start,end)
close_ss = datass["Close"]
close_jqr = datajqr["Close"]
close_ss.describe()
close_jqr.describe()

fig,ax = plt.subplots(nrows=1,ncols=2,figsize=(15,6))
close_ss.plot(ax=ax[0])
ax[0].set_title("SZZZ")
close_jqr.plot(ax=ax[1])
ax[1].set_title("JQR")


stock = pd.merge(datass,datajqr,left_index = True, right_index = True)
stock = stock[["Close_x","Close_y"]]
stock.columns = ["SZZZ","JQR"]
stock.head()

daily_return = (stock.diff()/stock.shift(periods = 1)).dropna()
daily_return.head()
daily_return.describe()

daily_return[daily_return["JQR"] > 0.105]

fig,ax = plt.subplots(nrows=1,ncols=2,figsize=(15,6))
daily_return["SZZZ"].plot(ax=ax[0])
ax[0].set_title("SZZZ")
daily_return["JQR"].plot(ax=ax[1])
ax[1].set_title("JQR")

fig,ax = plt.subplots(nrows=1,ncols=2,figsize=(15,6))
sns.distplot(daily_return["SZZZ"],ax=ax[0])
ax[0].set_title("SZZZ")
sns.distplot(daily_return["JQR"],ax=ax[1])
ax[1].set_title("JQR")

fig,ax = plt.subplots(nrows=1,ncols=1,figsize=(12,6))
plt.scatter(daily_return["JQR"],daily_return["SZZZ"])
plt.title("Scatter Plot of daily return between JQR and SZZZ")

daily_return["intercept"]=1.0
model = sm.OLS(daily_return["JQR"],daily_return[["SZZZ","intercept"]])
results = model.fit()
results.summary()


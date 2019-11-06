# -*- coding: utf-8 -*-
"""
Created on Sun Feb 26 22:26:55 2017

@author: Lenovo
"""

import itertools
import numpy as np
import pandas as pd

def mul_df(level1_rownum, level2_rownum, col_num):
    ''' create multilevel dataframe '''

    index_name = ['IDX_1','IDX_2']
    col_name = ['COL'+str(x).zfill(3) for x in range(col_num)]

    first_level_dt = [['A'+str(x).zfill(4)]*level2_rownum for x in range(level1_rownum)]
    first_level_dt = list(itertools.chain(*first_level_dt))
    second_level_dt = ['B'+str(x).zfill(3) for x in range(level2_rownum)]*level1_rownum

    dt = pd.DataFrame(np.random.randn(level1_rownum*level2_rownum, col_num), columns=col_name)
    dt[index_name[0]] = first_level_dt
    dt[index_name[1]] = second_level_dt

    rst = dt.set_index(index_name, drop=True, inplace=False)
    return rst

df = mul_df(4,5,3)

df.groupby(level='IDX_1').apply(lambda x: pd.rolling_sum(x,4))
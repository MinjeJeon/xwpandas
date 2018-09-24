import xlwings as xw
import xwpandas as xp
import pandas as pd
import numpy as np
import csv
from xlwings.constants import Constants as C
import importlib
import time

#%%

arrays = [np.array(['bar', 'bar', 'baz', 'baz', 'foo', 'foo', 'qux', 'qux']),
          np.array(['one', 'two', 'one', 'two', 'one', 'two', 'one', 'two'])]
index = pd.MultiIndex(levels=[['bar', 'baz', 'foo', 'qux'], [0, 'two']],
                      labels=[[0, 0, 1, 1, 2, 2, 3, 3], [0, 1, 0, 1, 0, 1, 0, 1]],
                      names=['first', 'second'])

mdf = pd.DataFrame(np.random.randn(6, 6), index=index[:6], columns=index[:6])
mdf.iloc[:, 2:4] = mdf.iloc[:, 2:4].astype(str).applymap(lambda x: x.replace('.','').replace('-',''))
bigmdf = pd.concat([mdf]*50000)

#%%
def time_elapsed(func, *args, **kwargs):
    import time
    start_time = time.time()
    res = func(*args, **kwargs)
    print("--- %s seconds ---" % round(time.time() - start_time, 2))
    return res
#%%
importlib.reload(xp.core)
time_elapsed(xp.save, bigmdf, '.temp/xl.xlsx')
time_elapsed(xp.core._df_toxlwings, bigmdf)
time_elapsed(lambda x:x.to_excel('.temp/xl.xlsx'), bigmdf)
def write_xlsxwriter(df, path):
    writer = pd.ExcelWriter(path, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1')
    writer.save()
    writer.close()
time_elapsed(write_xlsxwriter, bigmdf, '.temp/xl.xlsx')

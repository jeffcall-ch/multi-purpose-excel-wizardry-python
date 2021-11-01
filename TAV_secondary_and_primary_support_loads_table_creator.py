# This script filters the loads excel file extracted for Presenzano to those supports
# which are posing loads on supplier steel structure (with or without sec steel in supplier scope)
import pandas as pd
import numpy as np
from collections import namedtuple

# import TAV_support lists based on attached with secondary steel supplied by Enclosure supplier
# or supports attached directly to primary steel
# this results in 2 lists
TAV_support_xls = r"C:\Users\50000700\OneDrive - ansaldoenergiagroup\Desktop\Navis_xml\TAV Enclosure Supplier Input Supports.xlsx"
df_TAV_PRI_supp = pd.read_excel(TAV_support_xls, sheet_name='PRI_ALL', header=[0])
df_TAV_SEC_supp = pd.read_excel(TAV_support_xls, sheet_name='SEC_ALL', header=[0])
df_TAV_PRI_supp["System"] = df_TAV_PRI_supp["KKS"].str[:5]
df_TAV_SEC_supp["System"] = df_TAV_SEC_supp["KKS"].str[:5]
pri_supports = sorted(df_TAV_PRI_supp['KKS'].tolist())
sec_supports = sorted(df_TAV_SEC_supp['KKS'].tolist())

# read loads
loads_xls = r"C:\Users\50000700\PycharmProjects\multi-purpose-excel-wizardry-python\TAV_Support_Loads_from_PRE.xlsx"
df_loads = pd.read_excel(loads_xls, header=[1])
print(df_loads)
print (df_loads.size)
print(df_loads.keys())

# drop all duplicate KKS. Only one shall be present
df_loads = df_loads.drop_duplicates(subset=['Pipe Support KKS'])

# filter load input to only those KKS which are in scop of supplier (2x lists concat)
df_loads_filtered = df_loads[df_loads['Pipe Support KKS'].isin(sec_supports + pri_supports)]
# add "X" to supplier scope col
df_loads_filtered['Sec. steel in supplier scope'] = np.where(df_loads_filtered['Pipe Support KKS'].isin(sec_supports), "X", " ")
# group and sort
df_loads_filtered = df_loads_filtered.groupby('Sec. steel in supplier scope').apply(lambda x: x.sort_values('Pipe Support KKS'))
print(df_loads_filtered)
print (df_loads_filtered.size)

df_loads_filtered.to_excel("TAV_List_For_Supplier.xls")






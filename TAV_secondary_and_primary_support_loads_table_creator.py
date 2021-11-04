# This script filters the loads excel file extracted for Presenzano to those supports
# which are posing loads on supplier steel structure (with or without sec steel in supplier scope)
import openpyxl
import pandas as pd
import numpy as np
from collections import namedtuple

# import TAV_support lists based on attached with secondary steel supplied by Enclosure supplier
# or supports attached directly to primary steel
# this results in 2 lists
TAV_support_xls = r"C:\Users\50000700\OneDrive - ansaldoenergiagroup\Desktop\Navis_xml\TAV Enclosure Supplier Input Supports.xlsx"
df_TAV_PRI_supp = pd.read_excel(TAV_support_xls, sheet_name='PRI_ALL', header=[0])
df_TAV_SEC_supp = pd.read_excel(TAV_support_xls, sheet_name='SEC_ALL', header=[0])
df_TAV_SEC_and_PRI_supp = df_TAV_PRI_supp.append(df_TAV_SEC_supp)
# df_TAV_PRI_supp["System"] = df_TAV_PRI_supp["KKS"].str[:5]
# df_TAV_SEC_supp["System"] = df_TAV_SEC_supp["KKS"].str[:5]
pri_supports = sorted(df_TAV_PRI_supp['Pipe Support KKS'].tolist())
sec_supports = sorted(df_TAV_SEC_supp['Pipe Support KKS'].tolist())

# read loads
loads_xls = r"C:\Users\50000700\OneDrive - ansaldoenergiagroup\Desktop\Navis_xml\TAV_Support_Loads_from_PRE.xlsx"
df_loads = pd.read_excel(loads_xls, header=[1])
print(df_loads)
print (df_loads.size)
print(df_loads.keys())

# defining loads for not stress calculated supports
# below dict defines {nps:kN} values
not_stress_calc_loads = {0.5:-0.1, 1:-0.1, 1.5:-0.15, 2:-0.2, 3:-0.4, 4:-0.5}

# drop all duplicate KKS. Only one shall be present
df_loads = df_loads.drop_duplicates(subset=['Pipe Support KKS'])

# filter load input to only those KKS which are in scop of supplier (2x lists concat)
df_loads_filtered = df_loads[df_loads['Pipe Support KKS'].isin(sec_supports + pri_supports)]
# merge with df_TAV_SEC_and_PRI_supp in order to show group column
df_loads_filtered = df_loads_filtered.merge(df_TAV_SEC_and_PRI_supp, how='left')
# add "X" to supplier scope col
df_loads_filtered['Sec. steel in supplier scope'] = np.where(df_loads_filtered['Pipe Support KKS'].isin(sec_supports), "X", " ")
# group and sort
df_loads_filtered = df_loads_filtered.groupby('Sec. steel in supplier scope').apply(lambda x: x.sort_values('Pipe Support KKS'))
# write load value to supports which are not stress calculated


def set_kN(row):
    if (row['NPS Pipe'] in not_stress_calc_loads.keys()) and (row['-Z.1'] == "-"):
        return not_stress_calc_loads[row['NPS Pipe']]
    else:
        return row['-Z.1']


df_loads_filtered['-Z.1'] = df_loads_filtered.apply (lambda row: set_kN(row), axis=1)
df_loads_filtered = df_loads_filtered.sort_values(by=['Group'])

print(df_loads_filtered)
print (df_loads_filtered.size)

df_loads_filtered.to_excel(r"C:\Users\50000700\OneDrive - ansaldoenergiagroup\Desktop\Navis_xml\TAV_List_For_Supplier.xlsx")








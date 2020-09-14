import pandas as pd
import pathlib
import logging

df = pd.read_excel(r'C:\Users\50000700\Python\Python_repos\multi-purpose-excel-wizardry-python\excel_input\MBV-AD-numbers.xlsx')

# print (df)

df_formatted = pd.DataFrame({'AD':['FIRST AD'],
                   'KKS':['FIRST KKS']})

# print (df_formatted)

# Strip whitespace in all dataframe "cells"
df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)

for (index, row_series) in df.iterrows():
    if len(row_series) > 1:
        current_AD = row_series["AD"]
        #print (current_AD)
        #row_series.drop(labels=["AD"])
        for item in row_series.drop(labels=["AD"]):
            #print (item)
            df_current = pd.DataFrame({'AD':[current_AD],
                   'KKS':[item]})
            df_formatted = df_formatted.append(df_current)
    # print (row_series[0])

df_formatted = df_formatted.dropna()
print (df_formatted)

#writer = pd.ExcelWriter(r'C:\Users\50000700\Python\Python_repos\multi-purpose-excel-wizardry-python\excel_input\MBV-AD-numbers_formatted.xlsx')
#df_formatted.to_excel(writer, sheet_name='Sheet1')

df_formatted.to_excel(r'C:\Users\50000700\Python\Python_repos\multi-purpose-excel-wizardry-python\excel_input\MBV-AD-numbers_formatted.xlsx')

#writer.save
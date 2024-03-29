import pandas as pd
import pprint
import csv

rupsha_support_xls = r"C:\Users\50000700\PycharmProjects\multi-purpose-excel-wizardry-python\SupportList.xls"
df_supp = pd.read_excel(rupsha_support_xls, header=[0], sheet_name=1)
# print (len(df_supp))
df_supp['CHECK OK'] = df_supp['CHECK OK'].str.lower()
df_supp.columns = df_supp.columns.str.strip()
df_supp = df_supp.drop_duplicates(subset=['AD NUMBER'])

print (df_supp)
print (len(df_supp))

systems = df_supp['System'].unique()
print(systems)

# print(df_supp.keys())



all_total_dwgs = 0
all_ready_dwgs = 0

# comment added

summary_dicts = []
print (df_supp.loc[df_supp['System'] == "MBX42"])

for system in systems:
    # total_no_of_dwgs_per_system = df_supp.loc[df_supp.System == system, 'System'].count()
    total_no_of_dwgs_per_system = (df_supp.System == system).sum()
    ready_no_of_dwgs_per_system = len(df_supp[(df_supp['System']==system) & (df_supp['CHECK OK']=='x')])
    # ready_no_of_dwgs_per_system = df_supp[(df_supp['System']==system) & (df_supp['CHECK OK']=='x')].sum()
    print(f"system: {system} total in system: {total_no_of_dwgs_per_system} ready: {ready_no_of_dwgs_per_system}")
    all_total_dwgs += total_no_of_dwgs_per_system
    all_ready_dwgs += ready_no_of_dwgs_per_system
    current_row = {"System":system, "Ready":ready_no_of_dwgs_per_system, "All":total_no_of_dwgs_per_system}
    summary_dicts.append(current_row)

summary_dicts.append({"System":"Summary", "Ready":all_ready_dwgs, "All":all_total_dwgs})

df = pd.DataFrame(summary_dicts)
df = df.sort_values(by=['System'])

df.to_excel("STATUS.xls")

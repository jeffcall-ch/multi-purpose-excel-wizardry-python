import pandas as pd
import pprint
import csv

rupsha_PLL_xls = r"C:\Users\50000700\Python\Python_repos\multi-purpose-excel-wizardry-python\Rupsha_PLL_LS.xls"
df_pll = pd.read_excel(rupsha_PLL_xls, header=[0])

rupsha_mto_xls = r"C:\Users\50000700\Python\Python_repos\multi-purpose-excel-wizardry-python\piping_mto_.xls"
df_mto = pd.read_excel(rupsha_mto_xls, header=[0])
df_mto['DesAndPipeSpec'] = df_mto['Designation'] + df_mto['PipeSpec']

df_mto = df_mto.drop_duplicates(subset=['DesAndPipeSpec'])

# print (df_mto)

print (len(df_mto))


def getDictFromDf(df, kksColumnName, pSpecColumnName):
    list_of_KKS = df[kksColumnName].tolist()
    list_of_PipeSpec = df[pSpecColumnName].tolist()
    kks_pspec_dict = {}

    for i, kks in enumerate(list_of_KKS):
        kks_pspec_dict[kks] = list_of_PipeSpec[i]

    return kks_pspec_dict

pll_dict = getDictFromDf(df_pll, 'Designation', 'PipeSpec')
mto_dict = getDictFromDf(df_mto, 'Designation', 'PipeSpec')


# pprint.pprint (mto_dict)

kks_not_found_list = []
kks_pspec_mismatch_list = []

for kks in pll_dict.keys():
    if kks not in mto_dict.keys():
        kks_not_found_list.append(kks)
    else:
        if pll_dict[kks] != mto_dict[kks]:
            kks_pspec_mismatch_list.append(kks)

print(kks_not_found_list)
print(f"PLL kks not found in MTO list: {len(kks_not_found_list)}")
print(f"Psepc mismatch between PLL and MTO: {len(kks_pspec_mismatch_list)}")
print(kks_pspec_mismatch_list)


with open(r'C:\Users\50000700\Python\Python_repos\multi-purpose-excel-wizardry-python\mismatch.csv', 'w', newline='') as mismatchfile:
    writer = csv.writer(mismatchfile)
    writer.writerow(["No", "Designation", "PLL_PipeSpec", "S3D_PipeSpec"])
    for i, kks in enumerate(kks_pspec_mismatch_list):
        writer.writerow([i+1, kks, pll_dict[kks], mto_dict[kks]])
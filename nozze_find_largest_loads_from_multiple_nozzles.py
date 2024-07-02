import pandas as pd

results_xls = r"C:\R2NozzleCheck\results.xls"
results_pd = pd.read_excel(results_xls, header=[2])

results_pd = results_pd.fillna(method='ffill')


cols_to_drop = [27,26,25,24,22,21,20,18,17,16,14,13,12,10,9,8,6,5,4,2]
for i in cols_to_drop:
    results_pd = results_pd.drop(results_pd.columns[i], axis=1)

results_pd = results_pd.rename(columns={'Unnamed: 3': 'Fx', 'Unnamed: 7': 'Fy', 'Unnamed: 11': 'Fz', 'Unnamed: 15': 'Mx', 'Unnamed: 19': 'My', 'Unnamed: 23': 'Mz'})
results_pd = results_pd[results_pd["Loadcase"].str.contains('Extreme')]
results_pd = results_pd.drop(results_pd.columns[1], axis=1)

components = ['Fx', 'Fy', 'Fz', 'Mx', 'My', 'Mz']

for i in components:
    results_pd[i] = results_pd[i].abs()


Fax = []
Frad = []
Fz = []
Max = []
Mrad = []
Mz = []

df_list = results_pd.values.tolist()

for row in df_list:
    if 'X' in row[0]:
        Fax.append(row[1])
        Frad.append(row[2])
        Fz.append(row[3])
        Max.append(row[4])
        Mrad.append(row[5])
        Mz.append(row[6])

    if 'Y' in row[0]:
        Frad.append(row[1])
        Fax.append(row[2])
        Fz.append(row[3])
        Mrad.append(row[4])
        Max.append(row[5])
        Mz.append(row[6])

#print (Mrad)

Fax_max = max(Fax)
Frad_max = max(Frad)
Fz_max = max(Fz)

Max_max = max(Max)
Mrad_max = max(Mrad)
Mz_max = max(Mz)

print(f"Fax_max= {Fax_max} \t Frad_max= {Frad_max} \t Fz_max= {Fz_max}")
print(f"Max_max= {Max_max} \t Mrad_max= {Mrad_max} \t Mz_max= {Mz_max}")

#print(results_pd)
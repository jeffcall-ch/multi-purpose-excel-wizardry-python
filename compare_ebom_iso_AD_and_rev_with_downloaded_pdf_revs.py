# this script loads the list of AD numbers of an EBOM excel and compares the list
# with the downloaded pdf names

from pathlib import Path
import pandas as pd
import xlrd

# recursively find all xls files in a folder (and its subfolders)
FOLDER = r'C:\Users\50000700\OneDrive - ansaldoenergiagroup\Desktop\RUP isos\all isos'
folder = Path(FOLDER).glob('**/*.pdf')
listOfFiles = [x.stem[0:13] for x in folder if x.is_file()]
# print (listOfFiles)
# print (len(listOfFiles))


def create_ad_rev_dict(input_list):
    dict = {}
    for item in input_list:
        ad_no = item[0:10]
        rev = item[11:13]
        dict[str(ad_no)] = rev
    return dict
    

files_dict =create_ad_rev_dict(listOfFiles)
print (len(files_dict))
# print (files_dict)


# save the list of AD numbers into a blank excel sheet with only this info
isos_xls = r"C:\Users\50000700\OneDrive - ansaldoenergiagroup\Desktop\RUP isos\Iso_AD_with_rev.xls"
df_iso = pd.read_excel(isos_xls, header=[0])
# print (df_iso)
ebom_list = df_iso['AD Nos'].tolist()

ebom_dict = create_ad_rev_dict(ebom_list)
print (len(ebom_dict))
# print (ebom_dict)

mismatch_dict = {}
for item in ebom_dict.keys():
    if item in files_dict.keys():
        if ebom_dict[item] == files_dict[item]:
            continue
    mismatch_dict[item] = ebom_dict[item]
print ("MISMATCH DICT")
print (mismatch_dict)
print (len(mismatch_dict))



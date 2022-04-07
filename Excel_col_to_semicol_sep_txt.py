import pandas as pd

AE_iso_list_no_revisions_xls = r"C:\Users\50000700\OneDrive - ansaldoenergiagroup\Desktop\Weber documentation\Already approved supports list.xls"
AE_iso_list_no_revisions = pd.read_excel(AE_iso_list_no_revisions_xls, header=[0])
AE_iso_list_no_revisions = AE_iso_list_no_revisions.drop_duplicates(subset='AD', keep='first')
AE_iso_list_no_revisions = AE_iso_list_no_revisions.sort_values('AD')
AE_iso_list = AE_iso_list_no_revisions['AD'].tolist()

print(len(AE_iso_list))
iso_list_string_for_TC = ""

for item in AE_iso_list:
    iso_list_string_for_TC += item + ";"

text_file = open(r"C:\Users\50000700\OneDrive - ansaldoenergiagroup\Desktop\Weber documentation\iso_list_for_TX.txt", "w")
text_file.write(iso_list_string_for_TC)
text_file.close()


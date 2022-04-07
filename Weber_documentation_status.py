import pandas as pd
import pprint
import csv

# weber_isos_xls = r"C:\Users\50000700\OneDrive - ansaldoenergiagroup\Desktop\Weber documentation\Weber iso status with revisions 30.03.2022.xls"
weber_isos_xls = r"C:\Users\50000700\OneDrive - ansaldoenergiagroup\Desktop\Weber documentation\Weber supports status with revisions 30.03.2022.xls"
weber_isos = pd.read_excel(weber_isos_xls, header=[0])
# TC_isos_xls = r"C:\Users\50000700\OneDrive - ansaldoenergiagroup\Desktop\Weber documentation\Iso_list_WITH_revisions_TC_31.03.2022.xls"
TC_isos_xls = r"C:\Users\50000700\OneDrive - ansaldoenergiagroup\Desktop\Weber documentation\Support_list_WITH_LATEST_revisions_TC_USE_THIS_TO_COMPARE_MBH55_ADDED_TO_END_31.03.2022.xls"
TC_isos = pd.read_excel(TC_isos_xls, header=[0])

def write_list_of_iso_ADs_into_txt_file():
    AE_iso_list_no_revisions_xls = r"C:\Users\50000700\OneDrive - ansaldoenergiagroup\Desktop\Weber documentation\Support_list_with_revisions_from_EBOM_31.03.2022_formatted.xls"
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
    return iso_list_string_for_TC

# write_list_of_iso_ADs_into_txt_file()

weber_isos = weber_isos.drop_duplicates()
weber_isos = weber_isos.sort_values('AD_NUMBER')

#TC list contains all revisions, we keep only last duplicate values with the highest revision
TC_isos = TC_isos.drop_duplicates(['AD_NUMBER'], keep="last")

#print (weber_isos)
#print(TC_isos)

def get_dict_of_AD_REV_from_df(df):
    pairs = {}
    for index, row in df.iterrows():
        #print (row['AD_NUMBER'], row['REV'])
        pairs[row['AD_NUMBER']] = row['REV']
    #print (pairs)
    return pairs

weber_iso_ad_pairs = get_dict_of_AD_REV_from_df(weber_isos)
TC_iso_ad_pairs = get_dict_of_AD_REV_from_df(TC_isos)



def compare_dict (dict1, dict2):
    differences = []
    for key in dict1:
        dict1_key = key
        dict1_value = dict1[key]

        if dict1_key in dict2:
            dict2_value = dict2[dict1_key]
            if dict2_value != dict1_value:
                differences.append([dict1_key, dict1_value, dict2_value])
        else :
            differences.append([dict1_key, dict1_value, -1])
    return differences

#differences_TC_vs_Weber = compare_dict(TC_iso_ad_pairs, weber_iso_ad_pairs)
differences_TC_vs_Weber = pd.DataFrame(compare_dict(TC_iso_ad_pairs, weber_iso_ad_pairs), columns=['AD_NUMBER', 'REV_TC', 'REV_WEBER'])


print (differences_TC_vs_Weber)

#differences_Weber_vs_TC = compare_dict(weber_iso_ad_pairs, TC_iso_ad_pairs)
differences_Weber_vs_TC = pd.DataFrame(compare_dict(weber_iso_ad_pairs, TC_iso_ad_pairs), columns=['AD_NUMBER', 'REV_WEBER', 'REV_TC'])
print (differences_Weber_vs_TC)

differences_TC_vs_Weber.to_excel(r"C:\Users\50000700\OneDrive - ansaldoenergiagroup\Desktop\Weber documentation\TC_vs_WEBER.xls")
differences_Weber_vs_TC.to_excel(r"C:\Users\50000700\OneDrive - ansaldoenergiagroup\Desktop\Weber documentation\WEBER_vs_TC.xls")






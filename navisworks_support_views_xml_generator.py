# This script generates Navisworks xml input file for all pipe support locations
import pandas as pd
import numpy as np
from collections import namedtuple
from pprint import pprint

# INPUTS - DON'T FORGET TO SET THE PROJECT
input_xml = r"C:\Users\50000700\OneDrive - ansaldoenergiagroup\Desktop\Navis_xml\input.xml"
template_xml = r"C:\Users\50000700\OneDrive - ansaldoenergiagroup\Desktop\Navis_xml\template.xml"
PROJECT = "ZERO" # "ZERO", "MARGHERA", "PRESENZANO", "MINHANG", "TAVAZZANO", "FUSINA", "LA SPEZIA", "RUPSHA"

# Define rotation and shift of the coordinates - transformation
transformation_dict = {}
Transformation = namedtuple("Transformation", "alfa x y z")
transformation_dict["ZERO"] = Transformation(0,0,0,0)
transformation_dict["MARGHERA"] = Transformation(180, 113.9, 251.75, 6.6)
transformation_dict["PRESENZANO"] = Transformation(180, 180.497, 155.0, 6.6)
# transformation_dict["PRESENZANO"] = Transformation(180, 155.0, 180.497, 6.6)
transformation_dict["MINHANG"] = Transformation(0,0,0,0)
transformation_dict["TAVAZZANO"] = Transformation(0,0,0,0)
transformation_dict["FUSINA"] = Transformation(0,0,0,0)
transformation_dict["LA SPEZIA"] = Transformation(0,0,0,0)
transformation_dict["RUPSHA"] = Transformation(0,0,0,0)
transformation = transformation_dict[PROJECT]


def get_file_content_as_string(input_file_path):
    with open(input_file_path, 'r') as file:
        return file.read()


file_header = get_file_content_as_string(input_xml)
file_header = file_header[:(file_header.find("<viewpoints>")+17)]
template_string = get_file_content_as_string(template_xml)
# print(file_header)
# print(template_string)


class OneView:
    def __init__(self, kks_code, x, y, z, template_xml_string):
        self.x = x
        self.y = y
        self.z = z
        self.kks_code = kks_code
        self.template_xml_string = template_xml_string
        self.CAMERA_DISTANCE = 0
        self.CUT_PLANE_DIST = 2
        self.xml_variable_dictionary = {}
        self.xml_variable_dictionary["VIEW_NAME"] = self.kks_code
        self.xml_variable_dictionary["CAMPOS_X"] = self.x - self.CAMERA_DISTANCE
        self.xml_variable_dictionary["CAMPOS_Y"] = self.y - self.CAMERA_DISTANCE
        self.xml_variable_dictionary["CAMPOS_Z"] = self.z - self.CAMERA_DISTANCE
        self.xml_variable_dictionary["CLIP_MIN_X"] = self.x - self.CUT_PLANE_DIST
        self.xml_variable_dictionary["CLIP_MIN_Y"] = self.y - self.CUT_PLANE_DIST
        self.xml_variable_dictionary["CLIP_MIN_Z"] = self.z - self.CUT_PLANE_DIST
        self.xml_variable_dictionary["CLIP_MAX_X"] = self.x + self.CUT_PLANE_DIST
        self.xml_variable_dictionary["CLIP_MAX_Y"] = self.y + self.CUT_PLANE_DIST
        self.xml_variable_dictionary["CLIP_MAX_Z"] = self.z + self.CUT_PLANE_DIST

    def get_changed_xml_string(self):
        for k, v in self.xml_variable_dictionary.items():
            self.template_xml_string = self.template_xml_string.replace(k, str(v))
        return self.template_xml_string


# Support.xlsx processing
support_xls = r"C:\Users\50000700\OneDrive - ansaldoenergiagroup\Desktop\Navis_xml\Supports.xlsx"
df_supp = pd.read_excel(support_xls, header=[0])
alfa_rad = transformation.alfa * np.pi / 180
df_supp['Z'] = df_supp['Z'].div(1000).round(4) + transformation.z
df_supp['X_helper'] = (np.cos(alfa_rad)*df_supp['X'] - np.sin(alfa_rad)*df_supp['Y']).div(1000).round(5) + transformation.x
df_supp['Y_helper'] = (np.sin(alfa_rad)*df_supp['X'] + np.cos(alfa_rad)*df_supp['Y']).div(1000).round(5) + transformation.y

# put back values to the original columns
df_supp['X'] = df_supp['X_helper']
df_supp['Y'] = df_supp['Y_helper']
df_supp.drop('X_helper', inplace=True, axis=1)
df_supp.drop('Y_helper', inplace=True, axis=1)


df_supp.sort_values(by=['KKS'])
df_supp = df_supp.drop_duplicates(subset=['KKS'])
df_supp["System"] = df_supp["KKS"].str[:5]
print (df_supp)

# get all rows into list of lists
df_list = df_supp.values.tolist()
# print(df_list)

# get system abbreviations in ordered list
systems_list = df_supp["System"].unique()
systems_list = sorted(systems_list)
# print(systems_list)

# list of support rows grouped by systems
groupped_row_list = []
for system in systems_list:
    current_system_rows = []
    for row in df_list:
        if row[4] == system:
            current_system_rows.append(row)
    groupped_row_list.append(current_system_rows)

# pprint(groupped_row_list)

# Building the final xml string
final_xml_string = file_header
for row_list in groupped_row_list:
    folder_name = row_list[0][4] # get first element in list and the name of the system as folder name
    # prepend viewfolder tag with folder name
    final_xml_string = final_xml_string + '\n <viewfolder name="' + folder_name + '"> \n'
    for row in row_list:
        one_view = OneView(row[0],row[1],row[2],row[3], template_string)
        final_xml_string = final_xml_string + one_view.get_changed_xml_string()
    final_xml_string = final_xml_string + '\n </viewfolder> \n'

final_xml_string = final_xml_string + '\n </viewpoints> \n </exchange>'

# write xml final string into xml file
with open(r"C:\Users\50000700\OneDrive - ansaldoenergiagroup\Desktop\Navis_xml\output.xml", 'w') as file:
    file.write(final_xml_string)

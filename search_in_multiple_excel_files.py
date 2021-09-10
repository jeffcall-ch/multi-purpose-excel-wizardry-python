from pathlib import Path
import xlrd

# recursively find all xls files in a folder (and its subfolders)
FOLDER = r'C:\Users\50000700\Desktop\piping_classes'
#FOLDER = r'C:\Users\50000700\Desktop\Pipe Class Manuals'
folder = Path(FOLDER).glob('**/*.xls')
listOfFiles = [x for x in folder if x.is_file()]
# print (listOfFiles)

SEARCH_TEXT = "G2-D2"
found_list = []

def find_in_excel_file(input_file):
    sheet_data = []   
    wb = xlrd.open_workbook(file)
    sheets = wb.sheet_names()
    for sheet in sheets:
        sh = wb.sheet_by_name(sheet)
    for rownum in range(sh.nrows):
        sheet_data.append((sh.row_values(rownum)))
    for row in sheet_data:
        for cell in row:
            if SEARCH_TEXT in str(cell):
                # print(row)
                return(True)
    return(False)

for file in listOfFiles:
    if find_in_excel_file(file) == True:
        found_list.append(file)

print(f"SEARCH_TEXT: {SEARCH_TEXT}")
print("Files found with search string:")
print(found_list)

text_file = open(r'C:\Users\50000700\Desktop\piping_classes\SEARCH_OUTPUT.txt', "w")
text_file.write(str(found_list))
text_file.close()
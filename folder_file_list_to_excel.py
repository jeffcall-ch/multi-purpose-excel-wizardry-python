import pathlib
from pathlib import Path
import logging
import os
import xlwt
from tempfile import TemporaryFile


# find all files in all subfolders
# folder = Path(r'\\AENXSTOR01.aen.ansaldo.it\share08\Auxiliary_Baden\AES\GT36 S5\Marghera spring support check\New_Drawings\MBV groups').glob('**/*')

# find only spcific file type in current folder - NO subfolders are searched through
folder = Path(r'C:\Users\50000700\Desktop\TOTALS').glob('*.xls')

listOfFiles = [x.name for x in folder if x.is_file()]

# print (listOfFiles)

book = xlwt.Workbook()
sheet1 = book.add_sheet('sheet1')

for i,e in enumerate(listOfFiles):
    sheet1.write(i,1,e)

name = "list_of_files_in_folder.xls"
book.save(name)
book.save(TemporaryFile())


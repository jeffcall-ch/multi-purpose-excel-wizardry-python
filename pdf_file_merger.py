# merge pdf files
from  PyPDF2 import PdfFileMerger, PdfFileReader
from pathlib import Path

# recursively find all pdf files in a folder (and its subfolders)
# FOLDER = r'C:\Users\50000700\Desktop\MGL_Site_Visit_30.09.2021\Isometrics_MGL_22.09.2021'
FOLDER = r'C:\Users\50000700\Desktop\MGL_Site_Visit_30.09.2021\Pipe supports\Support dwgs per system'
folder = Path(FOLDER).glob('**/*.pdf')
listOfFiles = [x for x in folder if x.is_file()]
print (listOfFiles)
print (len(listOfFiles))

mergedObject = PdfFileMerger()

for file in listOfFiles:
    mergedObject.append(PdfFileReader(str(file)))

# mergedObject.write(r'C:\Users\50000700\Desktop\MGL_Site_Visit_30.09.2021\Isometrics_MGL_22.09.2021\all_isos.pdf')
mergedObject.write(r'C:\Users\50000700\Desktop\MGL_Site_Visit_30.09.2021\Pipe supports\Support dwgs per system\all_supports.pdf')
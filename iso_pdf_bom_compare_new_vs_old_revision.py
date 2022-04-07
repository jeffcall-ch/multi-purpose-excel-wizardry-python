# copy the BOM part of the particular page of new iso in one *.txt and the same from old into another *.txt
# script highlights differences
from rich import print

# INPUTS
new_iso_bom = r"C:\Users\50000700\OneDrive - ansaldoenergiagroup\Desktop\Iso text block check\new_iso_bom.txt"
old_iso_bom = r"C:\Users\50000700\OneDrive - ansaldoenergiagroup\Desktop\Iso text block check\old_iso_bom.txt"

with open(new_iso_bom, "r") as new_iso:
    lines = new_iso.readlines()

print(lines)
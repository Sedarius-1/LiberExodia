from __future__ import print_function
from PyPDF2 import PdfFileMerger, PdfMerger

import populate_library
from populate_library import *

workbook = populate_library.get_workbook("Templates/ExoditeUnits.xls")

for sheet in workbook.sheets():
    try:
        document = populate_library.get_template("Templates/HH Custom Unit Template - Left.docx")
        populate_unit_template(document, sheet)
    except (Exception,):
        print('Errors during creation of file')

workbook = populate_library.get_workbook("Templates/ExoditeWeapons.xls")

for sheet in workbook.sheets():
    try:
        document = populate_library.get_template("Templates/HH Weapons Template.docx")
        populate_weapons_template(document, sheet)
    except (Exception,):
        print('Errors during creation of file')

workbook = populate_library.get_workbook("Templates/ExoditeWargear.xls")

for sheet in workbook.sheets():
    try:
        document = populate_library.get_template("Templates/HH Wargear Template.docx")
        populate_wargear_template(document, sheet)
    except Exception as e:
        print(str(e))


merger = PdfMerger()

path_to_files = r'Unit_Cards/'
for root, dirs, file_names in os.walk(path_to_files):
    for file_name in file_names:
        if file_name != 'Wargear.pdf' and file_name != 'Weaponry.pdf':
            merger.append(path_to_files + file_name)
merger.append(path_to_files + 'Weaponry.pdf')
merger.append(path_to_files + 'Wargear.pdf')
merger.write(path_to_files+"Liber_Exodia.pdf")
merger.close()


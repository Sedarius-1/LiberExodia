from __future__ import print_function

import populate_library
from populate_library import *

if not os.path.exists('Unit_Cards'):
    os.mkdir('Unit_Cards')

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
#
workbook = populate_library.get_workbook("Templates/ExoditeWargear.xls")

for sheet in workbook.sheets():
    try:
        document = populate_library.get_template("Templates/HH Wargear Template.docx")
        populate_wargear_template(document, sheet)
    except Exception as e:
        print(str(e))

merger = PdfMerger()

path_to_files = r'Unit_Cards/ELITES'
append_to_pdf(path_to_files, merger)
path_to_files = r'Unit_Cards/HQ'
append_to_pdf(path_to_files, merger)
path_to_files = r'Unit_Cards/TROOPS'
append_to_pdf(path_to_files, merger)
path_to_files = r'Unit_Cards/FAST ATTACK'
append_to_pdf(path_to_files, merger)
path_to_files = r'Unit_Cards/LORD OF WAR'
append_to_pdf(path_to_files, merger)
path_to_files = r'Unit_Cards'
merger.append(path_to_files + '/Weaponry.pdf')
merger.append(path_to_files + '/Wargear.pdf')
merger.write(path_to_files + "/Liber_Exodia.pdf")
merger.close()

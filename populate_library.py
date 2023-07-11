import mailmerge
import xlrd
from docx2pdf import convert
import os
from comtypes import client


def get_workbook(filename: str) -> xlrd.book.Book:
    return xlrd.open_workbook(filename)


def get_template(template: str) -> mailmerge.MailMerge:
    return mailmerge.MailMerge(template)


def is_float(string_arg: str) -> bool:
    try:
        float(string_arg)
        return True
    except ValueError:
        return False


def float_string_to_string(float_string_arg: str) -> str:
    return str(int(float(float_string_arg)))


def get_current_key(current_index_arg: int, dictionary_arg: dict) -> str:
    return str(list(dictionary_arg.keys())[current_index_arg])


def get_current_value(current_index_arg: int, dictionary_arg: dict) -> str:
    return str(list(dictionary_arg.values())[current_index_arg])


def add_dict_entry(working_dictionary_arg: dict, xlsx_dictionary_arg: dict, current_index_arg: int) -> dict:
    key_elem = get_current_key(current_index_arg, xlsx_dictionary_arg)
    value_elem = get_current_value(current_index_arg, xlsx_dictionary_arg)
    working_dictionary_arg[key_elem[:len(key_elem) - 2]] = value_elem
    return working_dictionary_arg


def add_dict_entry_with_newlines(working_dictionary_arg: dict, xlsx_dictionary_arg: dict,
                                 current_index_arg: int) -> dict:
    key_elem = get_current_key(current_index_arg, xlsx_dictionary_arg)
    value_elem = get_current_value(current_index_arg, xlsx_dictionary_arg)
    value_elem = value_elem.replace('\\n', '\n')
    working_dictionary_arg[key_elem[:len(key_elem) - 2]] = value_elem
    return working_dictionary_arg


def add_empty_line_wargear(weapon_category_list_arg: list) -> None:
    weapon_category = dict()
    weapon_category['Weapon_Name'] = ' '
    weapon_category['Range'] = ' '
    weapon_category['Str'] = ' '
    weapon_category['AP'] = ' '
    weapon_category['Weapon_Type'] = ' '
    weapon_category_list_arg.append(weapon_category)


def add_header_line_wargear(weapon_category_list_arg: list) -> None:
    weapon_category = dict()
    weapon_category['Weapon_Name'] = 'Weapon Name'
    weapon_category['Range'] = 'Range'
    weapon_category['Str'] = 'Str'
    weapon_category['AP'] = 'AP'
    weapon_category['Weapon_Type'] = 'Weapon Type'
    weapon_category_list_arg.append(weapon_category)


def populate_unit_template(document_file: mailmerge, sheet_object: xlrd.book) -> None:
    col_a = sheet_object.col_values(0, 1)
    col_b = sheet_object.col_values(1, 1)

    imported_dictionary = {a: b for a, b in zip(col_a, col_b)}
    print(imported_dictionary)
    unit_member_stat_block = list()
    unit_members_amount_block = list()
    unit_members_type_block = list()
    wargear_block = list()
    special_rules_block = list()
    options_block = list()

    current_index = 0

    unit_name = get_current_value(current_index, imported_dictionary)
    current_index += 1

    points = float_string_to_string(get_current_value(current_index, imported_dictionary))
    current_index += 1

    unit_battlefield_role = get_current_value(current_index, imported_dictionary)
    current_index += 1

    unit_lore = ""

    while get_current_key(current_index, imported_dictionary)[
          :len(get_current_key(current_index, imported_dictionary)) - 2] == 'Unit_member_name':
        unit_member_stat = dict()
        for i in range(0, 11):
            key = get_current_key(current_index, imported_dictionary)
            if is_float(get_current_value(current_index, imported_dictionary)):
                value = float_string_to_string(get_current_value(current_index, imported_dictionary))
            else:
                value = get_current_value(current_index, imported_dictionary)
            unit_member_stat[key[:len(key) - 2]] = value
            current_index += 1
        unit_member_stat_block.append(unit_member_stat)

    while get_current_key(current_index, imported_dictionary)[
          :len(get_current_key(current_index, imported_dictionary)) - 2] == 'Unit_member':
        unit_members_amount = dict()
        unit_members_amount = add_dict_entry(unit_members_amount, imported_dictionary, current_index)
        current_index += 1
        unit_members_amount_block.append(unit_members_amount)

    while get_current_key(current_index, imported_dictionary)[
          :len(get_current_key(current_index, imported_dictionary)) - 2] == 'Unit_type':
        unit_members_type = dict()
        unit_members_type = add_dict_entry(unit_members_type, imported_dictionary, current_index)
        current_index += 1
        unit_members_type_block.append(unit_members_type)

    while get_current_key(current_index, imported_dictionary)[
          :len(get_current_key(current_index, imported_dictionary)) - 2] == 'Wargear_entry':
        wargear = dict()
        wargear = add_dict_entry(wargear, imported_dictionary, current_index)
        current_index += 1
        wargear_block.append(wargear)

    while get_current_key(current_index, imported_dictionary)[
          :len(get_current_key(current_index, imported_dictionary)) - 2] == 'Special_rule_entry':
        special_rules = dict()
        special_rules = add_dict_entry(special_rules, imported_dictionary, current_index)
        current_index += 1
        special_rules_block.append(special_rules)

    while get_current_key(current_index, imported_dictionary)[
          :len(get_current_key(current_index, imported_dictionary)) - 2] == 'Option_value':
        options = dict()
        options = add_dict_entry_with_newlines(options, imported_dictionary, current_index)
        current_index += 1
        options = add_dict_entry_with_newlines(options, imported_dictionary, current_index)
        current_index += 1

        options_block.append(options)

    document_file.merge(Unit_name=unit_name,
                        Points=points, Unit_battlefield_role=unit_battlefield_role, Unit_lore=unit_lore)

    document_file.merge_rows('Unit_member_name', unit_member_stat_block)
    document_file.merge_rows('Unit_member', unit_members_amount_block)
    document_file.merge_rows('Unit_type', unit_members_type_block)
    document_file.merge_rows('Wargear_entry', wargear_block)
    document_file.merge_rows('Special_rule_entry', special_rules_block)
    document_file.merge_rows('Option_value', options_block)

    document_file.write('Unit_Cards/' + unit_name + '.docx')
    convert('Unit_Cards/' + unit_name + '.docx', 'Unit_Cards/' + unit_name + '.pdf')
    os.remove('Unit_Cards/' + unit_name + '.docx')


def populate_weapons_template(document_file: mailmerge, sheet_object: xlrd.book) -> None:
    col_a = sheet_object.col_values(0, 1)
    col_b = sheet_object.col_values(1, 1)

    imported_dictionary = {a: b for a, b in zip(col_a, col_b)}
    # print(imported_dictionary)

    current_index = 0
    weapon_category_block = list()

    while (get_current_key(current_index, imported_dictionary)) != 'EOF':
        if (get_current_key(current_index, imported_dictionary))[
           :len(get_current_key(current_index, imported_dictionary)) - 2] == 'Weapon_Category':
            add_empty_line_wargear(weapon_category_block)
            weapon_category = dict()
            weapon_category['Weapon_Name'] = get_current_value(current_index, imported_dictionary)
            weapon_category['Range'] = ' '
            weapon_category['Str'] = ' '
            weapon_category['AP'] = ' '
            weapon_category['Weapon_Type'] = ' '
            weapon_category_block.append(weapon_category)
            add_empty_line_wargear(weapon_category_block)
            add_header_line_wargear(weapon_category_block)
            current_index += 1
        if (get_current_key(current_index, imported_dictionary))[
           :len(get_current_key(current_index, imported_dictionary)) - 2] == 'Weapon_Name':
            weapon_category = dict()
            weapon_category['Weapon_Name'] = get_current_value(current_index, imported_dictionary)
            current_index += 1
            weapon_category['Range'] = get_current_value(current_index, imported_dictionary)
            current_index += 1
            weapon_category['Str'] = get_current_value(current_index, imported_dictionary)
            current_index += 1
            weapon_category['AP'] = get_current_value(current_index, imported_dictionary)
            current_index += 1
            weapon_category['Weapon_Type'] = get_current_value(current_index, imported_dictionary)
            current_index += 1
            weapon_category_block.append(weapon_category)
        else:
            current_index += 1

    # print(weapon_category_block)

    document_file.merge_rows('Weapon_Name', weapon_category_block)

    document_file.write('Unit_Cards/Weaponry.docx')
    convert('Unit_Cards/Weaponry.docx', 'Unit_Cards/Weaponry.pdf')
    os.remove('Unit_Cards/Weaponry.docx')


def populate_wargear_template(document_file: mailmerge, sheet_object: xlrd.book) -> None:
    col_a = sheet_object.col_values(0, 1)
    col_b = sheet_object.col_values(1, 1)

    imported_dictionary = {a: b for a, b in zip(col_a, col_b)}

    current_index: int = 0
    wargear_entry_block = list()
    foreword: str = ''
    while (get_current_key(current_index, imported_dictionary)) != 'EOF':
        print(get_current_key(current_index, imported_dictionary))
        if get_current_key(current_index, imported_dictionary) == 'Foreword':
            foreword = get_current_value(current_index, imported_dictionary)
            current_index += 1
        else:
            wargear_entry: dict = dict()
            wargear_entry = add_dict_entry_with_newlines(wargear_entry, imported_dictionary, current_index)
            wargear_entry_block.append(wargear_entry)
            current_index += 1

    document_file.merge(Foreword=foreword)
    document_file.merge_rows('Wargear_Entry', wargear_entry_block)

    document_file.write('Unit_Cards/Wargear.docx')
    convert('Unit_Cards/Wargear.docx', 'Unit_Cards/Wargear.pdf')
    os.remove('Unit_Cards/Wargear.docx')
# coding=utf-8
import os
import sys
import openpyxl
import collections

import xmltodict


def read_excel(path, sheet_name='Sheet1', start_column=1, start_row=1, key_column=1, end_row=1, end_col=1,
               export_direct_to_res=False):
    """read the Excel file

    :param path: the Excel file path
    :type path: string

    :param sheet_name: the sheet name
    :type sheet_name: string

    :param start_column: start column index, start from 1
    :type start_column: int

    :param start_row: start row index, start from 1
    :type start_row: int

    :param key_column: the key name's column index
    :type key_column: int

    :param end_row: end row index
    :type end_row: int

    :param end_col: end column index
    :type end_col: int

    :param export_direct_to_res: whether need to export trans direct to your project's res value folder or not
    :type export_direct_to_res: bool
    """
    workbook = openpyxl.load_workbook(path)

    sheet = workbook[sheet_name]

    rows = sheet.max_row
    cols = sheet.max_column
    real_end_row = end_row + 1 if end_row != 1 else rows
    real_end_col = end_col + 1 if end_col != 1 else cols
    res_dict = {}
    key_list = []
    read_key = False
    for col in range(start_column, real_end_col):
        language = sheet.cell(1, col).value
        trans_list = []

        for row in range(start_row, real_end_row):
            # key's name, put it in key list
            if not read_key:
                key_name = sheet.cell(row, key_column).value
                key_name = key_name if key_name is not None else 'blank_key_name'
                key_list.append(key_name.replace(' ', '_').lower())

            # get the cell value
            cell_value = sheet.cell(row, col).value
            trans_list.append(cell_value if cell_value is not None else "blank_value")
        read_key = True
        res_dict[language] = trans_list

    write_result_to_xml(res_dict, key_list)
    for folder in values_folders:
        file_path = '%s%s/strings.xml' % (res_path, folder)
        print 'exporting %s' % folder
        export_to_xml(file_path, res_dict[folder], key_list)
    print('finish...')


def export_to_xml(file_path, trans_list, key_list):
    string_ordered_dict_list = []
    string_array_ordered_dict_list = []

    with open(file_path, 'r') as f:
        data_dict = xmltodict.parse(f.read())['resources']

        # list OrderedDict
        # key: @name value: #text
        string_ordered_dict_list = data_dict['string']
        # key: @name value: item(list[string])
        string_array_data = data_dict['string-array']
        if type(string_array_data) is not list:
            string_array_ordered_dict_list.append(string_array_data)

        # print string_ordered_dict_list
        # print string_array_ordered_dict_list

    os.rename(file_path, file_path.replace('strings.xml', 'strings-backup.xml'))
    with open(file_path, 'w+') as f:
        ordered_dict = collections.OrderedDict()

        for item in string_ordered_dict_list:
            key = item['@name']
            value = item['#text']
            ordered_dict[key] = value
        for index in range(len(key_list)):
            key = key_list[index]
            ordered_dict[key] = trans_list[index]
        print ordered_dict

        f.write('<resources>\n')
        # 处理string
        for key, value in ordered_dict.items():
            f.write('    ')
            f.write(formatter.format(name=key, str=value))
            f.write('\n')

        # 处理string-array
        for item in string_array_ordered_dict_list:
            name = item['@name']
            array = item['item']
            f.write('    <string-array name=\"%s\">\n' % name)
            for item_text in array:
                f.write('        <item>%s</item>\n' % item_text)
            f.write('    </string-array>\n')

        f.write('</resources>')


def write_result_to_xml(res_dict, key_list):
    with open('./result.xml', mode='w+') as res_file:
        for (key, value) in res_dict.items():
            language = key
            res_file.write('<!-- {lang} -->\n'.format(lang=language))
            for index in range(len(value)):
                lang_value = value[index].strip()
                if lang_value != 'blank_value':
                    if lang_value.startswith('<string'):
                        res_file.write(lang_value)
                    else:
                        res_file.write(formatter.format(name=key_list[index], str=lang_value))

                    res_file.write('\n')
            res_file.write('\n\n')


'''
{name} : the key name placeholder
{str} : the translation placeholder
'''
formatter = '<string name=\"{name}\">{str}</string>'
values_folders = ['values', 'values-zh-rCN']
res_path = './res/'
# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf8')
    read_excel(
        'trans-sample.xlsx',
        sheet_name='Sheet1',
        start_column=2,
        start_row=2,
        key_column=1,
        end_col=3,
        end_row=5,
        export_direct_to_res=False
    )

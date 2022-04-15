# coding=utf-8

import sys
import openpyxl


def read_excel(path, sheet_name='Sheet1', start_column=1, start_row=1, key_column=1, end_row=1, end_col=1):
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

    """
    workbook = openpyxl.load_workbook(path)

    sheet = workbook[sheet_name]

    rows = sheet.max_row
    cols = sheet.max_column
    real_end_row = end_row + 1 if end_row != 1 else rows
    real_end_col = end_col + 1 if end_col != 1 else cols
    res_dict = {}
    key_list = []
    for col in range(start_column, real_end_col):
        language = sheet.cell(1, col).value
        trans_list = []

        for row in range(start_row, real_end_row):
            # key's name, put it in key list
            key_name = sheet.cell(row, key_column).value
            key_name = key_name if key_name is not None else 'blank_key_name'
            key_list.append(key_name.replace(' ', '_').lower())
            # get the cell value
            cell_value = sheet.cell(row, col).value
            trans_list.append(cell_value if cell_value is not None else "blank_value")

        res_dict[language] = trans_list

    write_result_to_xml(res_dict, key_list)
    print('finish...')


def write_result_to_xml(res_dict, key_list):
    with open('./result.xml', mode='w+') as res_file:
        for (key, value) in res_dict.items():
            language = key
            res_file.write('<!-- {lang} -->\n'.format(lang=language))
            for index in range(len(value)):
                lang_value = value[index].strip()
                if lang_value != 'blank_value':
                    res_file.write(formatter.format(name=key_list[index], str=lang_value))
            res_file.write('\n\n')


'''
{name} : the key name placeholder
{str} : the translation placeholder
'''
formatter = '<string name=\"{name}\">{str}</string>\n'

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf8')
    read_excel(
        'trans-sample.xlsx',
        start_column=2,
        start_row=2,
        key_column=1,
        end_row=5,
        end_col=3
    )

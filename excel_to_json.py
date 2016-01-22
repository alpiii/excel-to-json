import xlrd
from collections import OrderedDict
import simplejson


def create_json(excel_file_name, sheet_index):
    """
    Reading excel file and creating the json file
    :param excel_file_name: Excel file name
    :param sheet_index: Sheet index number to read
    """

    workbook = xlrd.open_workbook(excel_file_name)
    sheet = workbook.sheet_by_index(sheet_index)
    data_list = []

    # starting from row 1, row 0 is header
    for row_num in range(1, sheet.nrows):
        data_row = OrderedDict()
        for col_num in range(0, sheet.ncols):
            data_row[sheet.cell(0, col_num).value] = \
                sheet.cell(row_num, col_num).value

        data_list.append(data_row)

    # Serializing
    j = simplejson.dumps(data_list)

    # Creating json file
    with open('data.json', 'w') as f:
        f.write(j)


if __name__ == '__main__':
    create_json('SampleData.xlsx', 0)

import sys
import os
import pandas as pd
import openpyxl as xl


def readExcel(input_dir):
    """
    read file excel and split
    Args:
        input_dir: excel director
    return:
        excel_data: excel data
    """
    work_book = xl.load_workbook(input_dir)
    excel_data = pd.read_excel(input_dir)
    excel_data.columns.values
    return excel_data


def getListHeader(excel_data):
    """
    read file excel and split
    Args:
        excel_data: excel data
    return:
        header_data: header data
    """
    header = list(excel_data.columns.values)
    return header


def getListDataColumn(excel_data, title_name):
    """
    read file excel and split
    Args:
        excel_data: excel_data
        title_name: title column
    return:
        filter_data: filter data
    """
    filter_data = []
    column_data = excel_data[title_name]
    for cell in column_data:
        if cell not in filter_data:
            filter_data.append(cell)
    return filter_data


def createExcelFile(output_dir, name, data_frame):
    file_name = os.path.join(output_dir, f'{name}.xlsx')
    data_frame.to_excel(file_name, name)


def splitTableData(excel_data, filter_data, filter_title, output_dir):
    """
    read file excel and split
    Args:
        excel_data: excel_data
        filter_data: filter column
    return:
        tables: tables data
    """
    rows = len(excel_data.axes[0])
    for filter in filter_data:
        tmp_table = []
        for i in range(0, rows):
            row_data = excel_data.iloc[i, :]
            if (row_data[filter_title] == filter):
                tmp_table.append(row_data)
        
        data_frame = pd.DataFrame(tmp_table)
        createExcelFile(output_dir, filter, data_frame)


def run(input_dir, output_dir):
    """
    read file excel and split
    Args:
        input_dir: excel director
    """
    excel_data = readExcel(input_dir)
    header = getListHeader(excel_data)
    filter_data = getListDataColumn(excel_data, header[2])
    splitTableData(excel_data, filter_data, header[2], output_dir)


if __name__ == "__main__":
    try:
        run(*sys.argv[1:])
    except Exception as exc:
        print(f"failed to run split excel ({exc.args}).")
        print(f'{exc}')

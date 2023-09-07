import os
import pandas as pd
from tqdm import tqdm


def readExcel(input_dir):
    """
    read file excel and split
    Args:
        input_dir: excel director
    return:
        excel_data: excel data
    """
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
    total = len(filter_data)
    progresBar = tqdm(range(total), desc="Create Fille...")
    file = 0
    for filter in filter_data:
        tmp_table = []
        for i in range(0, rows):
            row_data = excel_data.iloc[i, :]
            if (row_data[filter_title] == filter):
                tmp_table.append(row_data)

        data_frame = pd.DataFrame(tmp_table)
        createExcelFile(output_dir, filter, data_frame)
        file += 1
        progresBar.update()


def inputSplitParam():
    """
    ask the user to enter input
    """
    print("==========================")
    input_dir = input("Input (excel file): ")
    output_dir = input("Output (specify folder): ")
    active_sheet = input("Specify sheet active [0, 1, ...]: ")
    filter = input("Specify filter [0, 1, ...]: ")
    print("==========================")
    param = {
        'inputDir': input_dir,
        'outputDir': output_dir,
        'activeSheet': int(active_sheet),
        'filter': int(filter),
    }
    return param


def splitMode(input_dir, output_dir, filter):
    """
    read file excel and split
    Args:
        input_dir: excel director
    """
    excel_data = readExcel(input_dir)
    header = getListHeader(excel_data)
    filter_data = getListDataColumn(excel_data, header[filter])
    splitTableData(excel_data, filter_data, header[filter], output_dir)


def runSplitMode():
    """
    read file excel and split
    """
    param = inputSplitParam()
    input = param.get('inputDir')
    output = param.get('outputDir')
    filter = param.get('filter', 0)
    splitMode(input, output, filter)
    print("===========DONE===========")


def selectMode():
    """
    ask the user to specify mode
    """
    modes = ["Split file"]
    print("All mode:")
    i = 0
    for mode in modes:
        print(f'   {i}: {mode}')
        i += 1
    print("==========================")
    input_mode = input("select mode: ")
    return input_mode


def run(mode):
    """
    run specify mode
    Args:
        mode: selected mode
    """
    if (mode == 0):
        runSplitMode()
    else:
        print("mode dont exis")


if __name__ == "__main__":
    try:
        mode = selectMode()
        run(int(mode))
    except Exception as exc:
        print(f"failed to run split excel ({exc.args}).")
        print(f'{exc}')

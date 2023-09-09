# create by TurtleOfOceans 06/09/2023
import os
import openpyxl
import pandas as pd
from tqdm import tqdm


def readExcelWithPandas(input_dir, sheet=0):
    """
    read excel file with pandas lib
    Args:
        input_dir: excel director
        sheet: specify sheet to read data
    return:
        excel_data: excel data
    """
    excel_data = pd.read_excel(input_dir, sheet)
    excel_data.columns.values
    return excel_data


def readExcelWithOpenpyxl(input_dir, sheet=0):
    """
    read excel file with openpyxl lib
    Args:
        input_dir: excel director
        sheet: specify sheet to read data
    return:
        work_sheet: sheet data
    """
    work_book = openpyxl.load_workbook(input_dir, data_only=True)
    work_sheet = work_book.worksheets[sheet]
    return work_sheet


def getListHeader(excel_data):
    """
    get list header
    Args:
        excel_data: excel data
    return:
        header_data: header data
    """
    header = list(excel_data.columns.values)
    return header


def getListDataColumn(excel_data, title_name):
    """
    get list data in column
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


def createMultiExcelFile(output_dir, name, data_frame):
    """
    write data to excel file
    Args:
        output_dir: out put directory
        name: sheet name
        data_frame: data of sheet
    """
    file_name = os.path.join(output_dir, f'{name}.xlsx')
    data_frame.to_excel(file_name, name, index=False)


def splitTableDataToMultiFile(
        excel_data,
        filter_data,
        filter_title,
        output_dir):
    """
    split table data
    Args:
        excel_data: excel_data
        filter_data: filter column
        filter_title: specify column filter
        output_dir: specify directory out put
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
        createMultiExcelFile(output_dir, filter, data_frame)
        file += 1
        progresBar.update()


def splitTableDataToMultiSheet(
        excel_data,
        filter_data,
        filter_title,
        output_dir,
        sheet):
    """
    split table data
    Args:
        excel_data: excel_data
        filter_data: filter column
        filter_title: specify column filter
        output_dir: specify directory out put
    """
    rows = len(excel_data.axes[0])
    total = len(filter_data)
    file_name = os.path.join(output_dir, f'{sheet}.xlsx')
    writer = pd.ExcelWriter(file_name)
    progresBar = tqdm(range(total), desc="Create Fille...")
    file = 0
    excel_data.to_excel(writer, "Total", index=False)
    for filter in filter_data:
        tmp_table = []
        for i in range(0, rows):
            row_data = excel_data.iloc[i, 1:]
            if (row_data[filter_title] == filter):
                tmp_table.append(row_data)

        data_frame = pd.DataFrame(tmp_table)
        index = pd.Index(range(1, len(data_frame)+1))
        data_frame.index = index
        data_frame.to_excel(writer, filter, index_label="STT")
        file += 1
        progresBar.update()

    writer.close()


def inputSplitParam():
    """
    ask the user to enter input
    return:
        param: contain data specify by user
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


def inputParamForCreateFilter():
    """
    ask the user to enter input
    return:
        param: contain data specify by user
    """
    print("==========================")
    input_dir = input("Input (excel file): ")
    output_dir = input("Output (specify folder): ")
    active_sheet = input("Specify sheet active [0, 1, ...]: ")
    print("==========================")
    param = {
        'inputDir': input_dir,
        'outputDir': output_dir,
        'activeSheet': int(active_sheet),
    }
    return param


def splitDataWithMode(input_dir, output_dir, filter, sheet, mode):
    """
    read file excel and split
    Args:
        input_dir: excel director
        output_dir: specify out put data
        filter: specify filter
        sheet: specify sheet by user
        mode: specify mode by user
    """
    excel_data = readExcelWithPandas(input_dir, sheet)
    header = getListHeader(excel_data)
    filter_data = getListDataColumn(excel_data, header[filter])
    if (mode == 0):
        splitTableDataToMultiFile(
            excel_data,
            filter_data,
            header[filter],
            output_dir)
    elif (mode == 1):
        splitTableDataToMultiSheet(
            excel_data,
            filter_data,
            header[filter],
            output_dir,
            sheet)


def splitToMultiDataWithMode(mode):
    """
    read file excel and split
    Args:
        mode: split mode
    """
    param = inputSplitParam()
    input = param.get('inputDir')
    output = param.get('outputDir')
    filter = param.get('filter', 0)
    sheet = param.get('activeSheet', 0)
    splitDataWithMode(input, output, filter, sheet, mode)
    print("===========DONE===========")


def createFileWithOpenpyxl(rows_data, output_dir):
    """
    create file with openxlsx
    Args:
        output_dir: specify out put folder
        data: out put data
    """
    total = len(rows_data)
    print("Total: ", total)
    progresBar = tqdm(range(total), desc="Create Fille...")
    wb = openpyxl.Workbook()
    sheet = wb.active
    for i, row_data in enumerate(rows_data):
        for j in range(0, 13):
            sheet.cell(i + 1, j + 1).value = row_data[j].value
            progresBar.update()

    file_name = os.path.join(output_dir, f'Test.xlsx')
    wb.save(file_name)


def isRowNeed(row, min_clolumns, max_columns, row_number):
    """
    check row is copy
    Args:
        input_dir: specify input file
        sheet: active sheet
    """
    BG_COLOR = 64
    FG_COLOR = "FFFFFF00"
    isRow = True
    for i in range(min_clolumns, max_columns):
        cell = row[i]
        value = cell.value
        bg_color = cell.fill.bgColor.index
        fg_color = cell.fill.fgColor.index
        if value == None:
            continue
        if bg_color == BG_COLOR and fg_color == FG_COLOR:
            continue
        else:
            isRow = False
            break

    if isRow:
        for i in range(min_clolumns, max_columns):
            cell = row[i]
            if cell.value != None:
                bg_color = cell.fill.bgColor.index
                fg_color = cell.fill.fgColor.index
                if bg_color != BG_COLOR or fg_color != FG_COLOR:
                    isRow = False
                    break
    return isRow


def createTableDataWithColorFilter(input_dir, sheet, output_dir):
    """
    read and create filter color
    Args:
        input_dir: specify input file
        sheet: active sheet
        output_dir: specify output dir
    """
    rows_need = []
    work_sheet = readExcelWithOpenpyxl(input_dir, sheet)
    total = 0
    for i in range(2, work_sheet.max_row):
        row = work_sheet[i]
        isRow = isRowNeed(row, 0, 13, i)
        if isRow == True:
            # print(f'row[{i}] = {isRow}')
            total += 1
            rows_need.append(row)
    rows_length = len(rows_need)
    print(f'rows_length = {rows_length}, total = {total}')
    createFileWithOpenpyxl(rows_need, output_dir)


def createFilterColorMode(mode):
    """
    read and create filter color
    Args:
        mode: selected mode
    """
    param = inputParamForCreateFilter()
    input = param.get('inputDir')
    output = param.get('outputDir')
    sheet = param.get('activeSheet', 0)
    createTableDataWithColorFilter(input, sheet, output)
    print("===========DONE===========")


def selectMode():
    """
    ask the user to specify mode
    """
    modes = [
        "Split to multiple file",
        "Split to multiple sheet",
        "check color Mode"]
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
        splitToMultiDataWithMode(mode)
    elif (mode == 1):
        splitToMultiDataWithMode(mode)
    elif (mode == 2):
        createFilterColorMode(mode)
    else:
        print("mode dont exis")


if __name__ == "__main__":
    try:
        mode = selectMode()
        run(int(mode))
    except Exception as exc:
        print(f"failed to run split excel ({exc.args}).")
        print(f'{exc}')

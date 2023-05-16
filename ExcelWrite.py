import openpyxl
import parameters as ps
import datetime


def read_massive_from_excel(file_name=ps.data_excel,
                            min_row=ps.min_col,
                            max_row=ps.max_col,
                            min_col=ps.min_row,
                            max_col=ps.max_row):
    book = openpyxl.load_workbook(file_name)
    sheet = book.worksheets[0]

    # print(tuple(sheet.rows))
    reference_uniq_cols = []

    for col in sheet.iter_cols(min_row, max_row, min_col, max_col):
        help_list = []
        for cell in col:
            if cell.value not in help_list:
                help_list.append(cell.value)
        reference_uniq_cols.append(help_list)

    return reference_uniq_cols


def write_massive_to_excel(massive, col=1, row=1):
    wb = openpyxl.Workbook()
    ws = wb.active

    if type(massive) == list:
        for subarray in massive:
            for index, value in enumerate(subarray):
                ws.cell(column=col + index, row=row).value = value
            row += 1
    elif type(massive) == dict:
        for subarray in massive:
            ws.cell(column=col, row=row).value = subarray
            for index, value in enumerate(massive.get(subarray)):
                ws.cell(column=col, row=row + 1 + index).value = value
            col += 1

    dt = datetime.datetime.now()
    time = dt.time()
    now = str(time.strftime('%H ч %M мин'))

    wb.save(f"files/{now} " + ("Атрибуты" if type(massive) == dict else "Значения") + ".xlsx")

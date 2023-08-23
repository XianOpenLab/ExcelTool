import openpyxl
import pandas as pd
import xlrd


# 读取excel函数
def read_excel(path, sn=None):
    if path.endswith(".xlsx"):
        sheet_names = pd.ExcelFile(path).sheet_names
        if sn is None:
            # 读取.xlsx文件中的第二张表
            if "结课表" in sheet_names:
                df = pd.read_excel(path, sheet_name="结课表")
            else:
                df = pd.read_excel(path, sheet_name=sheet_names[0])
        else:
            df = pd.read_excel(path, sheet_name=sn)

        # 将所有数据放到二维列表中

        data: list = df.values.tolist()
        data.insert(0, df.keys().tolist())
        # 打印二维列表的内容

        return data

    else:
        file = xlrd.open_workbook(path)
        sheet = file.sheet_by_index(0)
        data = []
        for row in range(sheet.nrows):
            row_data = []
            for col in range(sheet.ncols):
                cell_value = sheet.cell_value(row, col)
                row_data.append(cell_value)
            data.append(row_data)
        return data


def get_row_len(row):
    row_len = 0
    for r in row:
        if r.value is None:
            break
        else:
            row_len += 1
    return row_len


def is_none(item):
    for value in item:
        if value is not None:
            return False
    return True


# 写入excel函数
def write_excel(path, row_data):
    wb = openpyxl.Workbook()
    sheet = wb.active  # 获取当前的表单
    for row in row_data:
        sheet.append(row)
    wb.save(path)
    print("写入数据成功！文件保存在：" + path)

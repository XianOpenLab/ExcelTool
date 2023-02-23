import openpyxl
import xlrd


# 读取excel函数
def read_excel(path):
    if path.endswith(".xlsx"):
        wb = openpyxl.load_workbook(path)  # 读取excel文件
        sheet = wb.active
        row_data = []
        row_len = 0
        for row in sheet.rows:
            if row_len == 0:
                row_len = get_row_len(row)
            item = [row[i].value for i in range(row_len)]
            if is_none(item):
                break
            else:
                row_data.append(item)
        return row_data
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


if __name__ == "__main__":
    __filePath = "/Users/musicbear/欧朋/test/2022.12.31-PYTHON-寒假班结课统计表.xlsx"
    read_excel(__filePath)

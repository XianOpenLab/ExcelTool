import os
import sys
import CsvParser
import ExcelParser
import warnings
import shutil
import pic2excel
from Sheet import Sheet
from tkinter import Tk
from tkinter import filedialog
from constants import *

warnings.filterwarnings('ignore')

WORK = 1
EXAM = 2
EXPORT_WITH_FORM = 1
EXPORT_NEW = 2
root = Tk()
root.wm_attributes('-topmost', 1)


def add_submit_count(excel, csv, export_type):
    for i in range(len(csv.data)):
        index = find(excel, csv.get(i, FIELD_NAME))
        if index is not None:
            if csv.get(i, FIELD_COMMIT_STATE) == "已提交":
                if export_type == EXPORT_NEW:
                    excel.set(index, csv.filename, "✓")
                else:
                    rate = excel.get(index, FIELD_WORK_COMPLETION)
                    excel.set(index, FIELD_WORK_COMPLETION, 1 if rate is None else rate + 1)
            elif export_type == EXPORT_NEW:
                excel.set(index, csv.filename, "×")


def calculate(excel, count, export_type):
    for i in range(len(excel.data)):
        student_name = excel.get(i, FIELD_NAME)
        if student_name is not None:
            if export_type == EXPORT_NEW:
                submit_count = excel.data[i].count("✓")
            else:
                submit_count = excel.get(i, FIELD_WORK_COMPLETION)
            rate_str = "%.2f" % (0 if submit_count is None else submit_count / count * 100) + "%"
            excel.set(i, FIELD_WORK_COMPLETION, rate_str)
            print(student_name + "作业完成率：" + rate_str)


def statistics_homework(excel, csv_list, export_type):
    count = 0
    excel.excel_init(FIELD_WORK_COMPLETION)
    for csv in csv_list:
        if not csv.title.__contains__(FIELD_COMMIT_STATE):
            print("《" + csv.filename + "》非作业文件，不对其进行统计")
            continue
        else:
            count += 1
            add_submit_count(excel, csv, export_type)
    if count > 0:
        calculate(excel, count, export_type)
        if export_type == EXPORT_NEW:
            output(excel, None)
        else:
            output(excel, excel.filename)


def is_float(string):
    if string is None:
        return False
    try:
        float(string)
        return True
    except ValueError:
        return False


def statistics_exam(excel, csv_list, export_type):
    excel.excel_init(FIELD_SCORE)
    for csv in csv_list:
        if not csv.title_contain(FIELD_SCORE):
            print("《" + csv.filename + "》非成绩单，不对其进行统计")
            continue
        else:
            if export_type == EXPORT_NEW:
                title = [FIELD_UID, FIELD_NAME, FIELD_OBJECTIVE_SCORE, FIELD_SUBJECTIVE_SCORE, FIELD_SCORE]
                form = []
                for i in range(len(csv.data)):
                    row = [csv.get(i, FIELD_UID), csv.get(i, FIELD_REAL_NAME), csv.get(i, FIELD_OBJECTIVE_SCORE),
                           csv.get(i, FIELD_SUBJECTIVE_SCORE), csv.get(i, FIELD_SCORE)]
                    form.append(row)
                form.insert(0, title)
                excel = Sheet("output.xlsx", form)
                output(excel, csv.filename)
            else:
                for i in range(len(csv.data)):
                    index = find(excel, csv.get(i, FIELD_REAL_NAME))
                    if index is not None:
                        excel.set(index, FIELD_SCORE, csv.get(i, FIELD_SCORE))
                output(excel, excel.filename)


def remove(excel, csv_list):
    for csv in csv_list:
        for i in range(len(csv.data)):
            index = find(excel, csv.get(i, FIELD_NAME))
            if index is not None:
                print("删除学员：" + csv.get(i, FIELD_NAME))
                del excel.data[index]
    return excel


def remain(excel, csv_list):
    remain_list = []
    new_excel = []
    for csv in csv_list:
        for i in range(len(csv.data)):
            name = csv.get(i, FIELD_NAME)
            index = find(excel, name)
            if index is not None and not remain_list.__contains__(name):
                new_excel.append(excel.data[index])
            if not remain_list.__contains__(name):
                remain_list.append(name)
    for i in range(len(excel.data)):
        name = excel.get(i, FIELD_NAME)
        if name is not None and not remain_list.__contains__(name):
            print("删除学员:" + name)
    new_excel.sort(key=index_sort)
    new_excel.insert(0, excel.title)
    return Sheet(excel.filename, new_excel)


def index_sort(item):
    return int(item[0])


def get_out_path(file_name):
    index = 0
    while True:
        if file_name is None:
            name = current_path + "out/output" + ("" if index == 0 else "(%d)" % index) + ".xlsx"
        else:
            name = current_path + "out/" + file_name + ("" if index == 0 else "(%d)" % index) + ".xlsx"
        if not os.path.exists(name):
            return name
        index += 1


def output(excel, file_name):
    if not os.path.exists(current_path + "out"):
        os.makedirs(current_path + "out")
    if file_name is not None:
        file_name = file_name.replace(".csv", "").replace(".xlsx", "").replace(".xls", "")
        ExcelParser.write_excel(get_out_path(file_name), excel.get_form_data())
    else:
        ExcelParser.write_excel(get_out_path(None), excel.get_form_data())


def find(excel, name):
    for i in range(len(excel.data)):
        if excel.get(i, FIELD_NAME) == name:
            return i
        elif excel.get(i, FIELD_REAL_NAME) == name:
            return i


def get_csv(step_type):
    if step_type == WORK:
        print("""
***********************************
选取进行统计的作业表(多选),按Enter继续
***********************************
                    """)
        input()
        files = select_files()
    else:
        print("""
**********************
选取成绩单表,按Enter继续
**********************
                            """)
        input()
        files = [select_file()]
    if files == ['']:
        return None
    csv_list = [Sheet(os.path.basename(csv),
                      CsvParser.read_csv(csv)
                      if csv.endswith(".csv") else
                      ExcelParser.read_excel(csv))
                for csv in files]
    return csv_list


def select_out_type(step_type):
    csv_list = get_csv(step_type)
    if csv_list is None or len(csv_list) == 0:
        return None, None, None
    print(
        """
******************************
按1，将结果导出到一个指定的表格中
按2，将结果导出到新的文件中
******************************
        """
    )
    while True:
        input_str = input("请输入:")
        input_str = input_str.strip()
        if input_str.isdigit():
            num = int(input_str)
            if num == EXPORT_WITH_FORM:
                file = select_file()
                if file is None or file == '':
                    return None, None, None
                elif file.endswith(".xlsx") or file.endswith(".xls"):
                    return Sheet(os.path.basename(file),
                                 ExcelParser.read_excel(file)), csv_list, EXPORT_WITH_FORM
                else:
                    return Sheet(os.path.basename(file), CsvParser.read_csv(file)), csv_list, EXPORT_WITH_FORM
            elif num == EXPORT_NEW:
                title = [FIELD_INDEX, FIELD_NAME]
                excel = Sheet("output.xlsx", [title])
                for csv in csv_list:
                    for i in range(len(csv.data)):
                        if find(excel, csv.get(i, FIELD_NAME)) is None:
                            e_row = [csv.get(i, FIELD_INDEX), csv.get(i, FIELD_NAME)]
                            excel.data.append(e_row)
                return excel, csv_list, EXPORT_NEW
            else:
                print("输入不正确")


def main_func(select_str):
    print("开始执行。。。")
    if select_str == "1":
        excel, csv_list, export_type = select_out_type(WORK)
        if excel is None or csv_list is None or export_type is None:
            return
        statistics_homework(excel, csv_list, export_type)
    elif select_str == "2":
        excel, csv_list, export_type = select_out_type(EXAM)
        if excel is None or csv_list is None or export_type is None:
            return
        statistics_exam(excel, csv_list, export_type)
    elif select_str == "3":
        if not delete_student():
            return
    elif select_str == "4":
        png = select_file(filetypes=[('pic', '*.png'), ('pic', '*.jpg'), ('pic', '*.jpeg')])
        if png is None or png == '':
            return
        file_path = pic2excel.image2excel(png)
        file_path = shutil.move(file_path, current_path + os.path.basename(file_path))
        print("excel文件已生成:" + file_path)
    elif select_str == "0":
        print("程序已退出")
        sys.exit()
    else:
        print("输入不正确，请重新输入")
    print("执行完毕")
    input("按Enter键继续。。。")


def delete_student():
    print("""
**********************************
指定要批量删除的"目标"表，按Enter键继续
**********************************
            """)
    input()
    excel = select_file()
    if excel == '':
        return False
    if excel.endswith(".xlsx") or excel.endswith(".xls"):
        e = ExcelParser.read_excel(excel)
    else:
        e = CsvParser.read_csv(excel)
    excel = Sheet(os.path.basename(excel), e)
    print("""
*****************************
指定"参考"表(多选)，按Enter键继续
*****************************
            """)
    input()
    csv_list = select_files()
    if csv_list == '':
        return False
    csv_list = [Sheet(os.path.basename(csv),
                      CsvParser.read_csv(csv)
                      if csv.endswith(".csv") else
                      ExcelParser.read_excel(csv))
                for csv in csv_list]
    print("""
*************************************
按1，"目标"将删除"参考"中存在的学员
按2，"目标"将保留"参考"中存在的学员其余删除
*************************************
            """)
    while True:
        num = input("请输入:")
        if num.isdigit():
            num = int(num)
            if num == 1:
                excel = remove(excel, csv_list)
                output(excel, excel.filename)
                break
            elif num == 2:
                excel = remain(excel, csv_list)
                output(excel, excel.filename)
                break
            else:
                print("未找到对应项")
        else:
            print("输入不正确")
    return True


def select_files(filetypes=None):
    if filetypes is None:
        filetypes = [('excel', '*.csv'), ('excel', '*.xlsx'), ('excel', '*.xls')]
    root.withdraw()
    file_list = filedialog.askopenfilenames(filetypes=filetypes, parent=root)
    root.update()
    return file_list


def select_file(filetypes=None):
    if filetypes is None:
        filetypes = [('excel', '*.csv'), ('excel', '*.xlsx'), ('excel', '*.xls')]
    root.withdraw()
    file = filedialog.askopenfilename(filetypes=filetypes, parent=root)
    root.update()
    return file


if __name__ == '__main__':
    current_path = os.path.realpath(os.path.dirname(sys.argv[0])) + "/"
    if not os.path.exists(current_path + "out"):
        os.makedirs(current_path + "out")

    m_menu = """
***********************
按1，统计作业提交率
按2，提取考试成绩
按3，批量删除学员
按4，图片转为excel文件
按0，退出程序
***********************"""
    print(m_menu)
    while True:
        in_s = input("请输入:").strip()
        if in_s.isdigit():
            # main_func(m_csv_list, in_s)
            main_func(in_s)
            print(m_menu)
        else:
            print("未找到对应项")

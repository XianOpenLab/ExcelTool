import os
import sys
import CsvParser
import ExcelParser
import warnings
import shutil
import pic2excel
from DataInstance import DataInstance
from tkinter import Tk
from tkinter import filedialog

warnings.filterwarnings('ignore')

WORK = 1
EXAM = 2
EXPORT_WITH_FORM = 1
EXPORT_NEW = 2


def add_submit_count(excel, csv, export_type):
    for i in range(len(csv.data)):
        index = find(excel, csv.get(i, "姓名"))
        if index is not None:
            if csv.get(i, "提交状态") == "已提交":
                if export_type == 2:
                    excel.set(index, csv.filename, "✓")
                else:
                    rate = excel.get(index, "作业完成率%")
                    excel.set(index, "作业完成率%", 1 if rate is None else rate + 1)


def calculate(excel, count, export_type):
    for i in range(len(excel.data)):
        student_name = excel.get(i, "姓名")
        if student_name is not None:
            if export_type == 2:
                submit_count = excel.data[i].count("✓")
            else:
                submit_count = excel.get(i, "作业完成率%")
            rate_str = "%.2f" % (0 if submit_count is None else submit_count / count * 100) + "%"
            excel.set(i, "作业完成率%", rate_str)
            print(student_name + "作业完成率：" + rate_str)


def statistics_homework(excel, csv_list, export_type):
    count = 0
    excel.excel_init("作业完成率%")
    for csv in csv_list:
        if not csv.title.__contains__("提交状态"):
            print("《" + csv.filename + "》非作业文件，不对其进行统计")
            continue
        else:
            count += 1
            add_submit_count(excel, csv, export_type)
    if count > 0:
        calculate(excel, count, export_type)
        if export_type == 2:
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


def get_score(csv, i):
    kg = csv.get(i, "客观题得分")
    zg = csv.get(i, "主观题得分")
    if is_float(kg) and is_float(zg):
        score = float(kg) + float(zg)
    elif is_float(kg):
        score = float(kg)
    else:
        score = float(zg)
    return score


def statistics_exam(excel, csv_list, export_type):
    excel.excel_init("考试成绩")
    for csv in csv_list:
        if not csv.title_contain("得分"):
            print("《" + csv.filename + "》非成绩单，不对其进行统计")
            continue
        else:
            if export_type == EXPORT_NEW:
                title = ["学员id", "姓名", "客观题得分", "主观题得分", "考试成绩"]
                form = []
                for i in range(len(csv.data)):
                    row = [csv.get(i, "学员id"), csv.get(i, "真实姓名"), csv.get(i, "客观题得分"),
                           csv.get(i, "主观题得分"), get_score(csv, i)]
                    form.append(row)
                form.insert(0, title)
                excel = DataInstance("output.xlsx", form)
                output(excel, csv.filename)
            else:
                for i in range(len(csv.data)):
                    index = find(excel, csv.get(i, "真实姓名"))
                    if index is not None:
                        excel.set(index, "考试成绩", get_score(csv, i))
                output(excel, excel.filename)


def remove(excel, csv_list):
    for csv in csv_list:
        for i in range(len(csv.data)):
            index = find(excel, csv.get(i, "姓名"))
            if index is not None:
                print("移除学员：" + csv.get(i, "姓名"))
                del excel.data[index]
    return excel


def remain(excel, csv_list):
    remain_list = []
    new_excel = []
    for csv in csv_list:
        for i in range(len(csv.data)):
            name = csv.get(i, "姓名")
            index = find(excel, name)
            if index is not None and not remain_list.__contains__(name):
                new_excel.append(excel.data[index])
            if not remain_list.__contains__(name):
                remain_list.append(name)
    for i in range(len(excel.data)):
        name = excel.get(i, "姓名")
        if name is not None and not remain_list.__contains__(name):
            print("移除学员:" + name)
    new_excel.sort(key=index_sort)
    new_excel.insert(0, excel.title)
    return DataInstance(excel.filename, new_excel)


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
        if excel.get(i, "姓名") == name:
            return i
        elif excel.get(i, "真实姓名") == name:
            return i


def get_csv(step_type):
    Tk().withdraw()
    if step_type == WORK:
        files = filedialog.askopenfilenames(filetypes=[('excel', '*.csv'), ('excel', '*.xlsx'), ('excel', '*.xls')])
    else:
        files = [filedialog.askopenfilename(filetypes=[('excel', '*.csv'), ('excel', '*.xlsx'), ('excel', '*.xls')])]
    csv_list = [DataInstance(os.path.basename(csv),
                             CsvParser.read_csv(csv)
                             if csv.endswith(".csv") else
                             ExcelParser.read_excel(csv))
                for csv in files]
    return csv_list


def select_out_type(step_type):
    csv_list = get_csv(step_type)
    print(
        """
*********************************
按1，将结果导出到一个指定的表格中
按2，将结果导出到新的文件中
*********************************
        """
    )
    while True:
        input_str = input("请输入:")
        input_str = input_str.strip()
        if input_str.isdigit():
            num = int(input_str)
            if num == EXPORT_WITH_FORM:
                Tk().withdraw()
                file = filedialog.askopenfilename(
                    filetypes=[('excel', '*.csv'), ('excel', '*.xlsx'), ('excel', '*.xls')])
                if file.endswith(".xlsx") or file.endswith(".xls"):
                    return DataInstance(os.path.basename(file),
                                        ExcelParser.read_excel(file)), csv_list, EXPORT_WITH_FORM
                else:
                    return DataInstance(os.path.basename(file), CsvParser.read_csv(file)), csv_list, EXPORT_WITH_FORM
            elif num == EXPORT_NEW:
                title = ['序号', '姓名']
                excel = DataInstance("output.xlsx", [title])
                for csv in csv_list:
                    for i in range(len(csv.data)):
                        if find(excel, csv.get(i, '姓名')) is None:
                            e_row = [csv.get(i, '序号'), csv.get(i, '姓名')]
                            excel.data.append(e_row)
                return excel, csv_list, EXPORT_NEW
            else:
                print("输入不正确")


def main_func(select_str):
    print("开始执行。。。")
    if select_str == "1":
        excel, csv_list, export_type = select_out_type(WORK)
        statistics_homework(excel, csv_list, export_type)
    elif select_str == "2":
        excel, csv_list, export_type = select_out_type(EXAM)
        statistics_exam(excel, csv_list, export_type)
    elif select_str == "3":
        delete_student()
    elif select_str == "4":
        Tk().withdraw()
        png = filedialog.askopenfilename(title="选择表格图片",
                                         filetypes=[('excel', '*.png'), ('excel', '*.jpg'), ('excel', '*.jpeg')])
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
*******************************
指定要移除学员的表格，按Enter键继续
*******************************
            """)
    input()
    Tk().withdraw()
    excel = filedialog.askopenfilename(title="选择待处理表格",
                                       filetypes=[('excel', '*.csv'), ('excel', '*.xlsx'), ('excel', '*.xls')])
    if excel.endswith(".xlsx") or excel.endswith(".xls"):
        e = ExcelParser.read_excel(excel)
    else:
        e = CsvParser.read_csv(excel)
    excel = DataInstance(os.path.basename(excel), e)
    print("""
*****************************
指定参考表格(多选)，按Enter键继续
*****************************
            """)
    input()
    Tk().withdraw()
    csv_list = filedialog.askopenfilenames(title="选择参考表格",
                                           filetypes=[('excel', '*.csv'), ('excel', '*.xlsx'), ('excel', '*.xls')])
    csv_list = [DataInstance(os.path.basename(csv),
                             CsvParser.read_csv(csv)
                             if csv.endswith(".csv") else
                             ExcelParser.read_excel(csv))
                for csv in csv_list]
    print("""
*************************************
按1，将"参考表格"中存在的学员移除
按2，将"参考表格"中存在的学员保留其他移出
*************************************
            """)
    num = input("请输入:")
    if num.isdigit():
        num = int(num)
        if num == 1:
            excel = remove(excel, csv_list)
        elif num == 2:
            excel = remain(excel, csv_list)
        output(excel, excel.filename)


if __name__ == '__main__':
    current_path = os.path.realpath(os.path.dirname(sys.argv[0])) + "/"
    if not os.path.exists(current_path + "out"):
        os.makedirs(current_path + "out")

    m_menu = """
***********************
按1，统计作业提交率(多选)
按2，提取考试成绩
按3，移除学员
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

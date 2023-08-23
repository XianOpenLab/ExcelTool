import os
import shutil
import sys
import traceback

from openlabTool import excelParser, csvParser, pic2excel
from openlabTool.constants import *
from openlabTool.customException import CsvReadException
from openlabTool.deleteInfo import DeleteInfo
from openlabTool.examinfo import ExamInfo
from openlabTool.student import Student
from openlabTool.workinfo import WorkInfo

work_path = "作业/"
exam_path = "考试/"
output = "结果/"
png_path = "图片/"
delete_path = "删除/"


def getWorkCommit():
    wi_ls = []
    for file in os.listdir(work_path):
        if file.endswith(".xlsx") or file.endswith(".xls"):
            wi_ls.append(WorkInfo(excelParser.read_excel(work_path + file)))
        elif file.endswith(".csv"):
            wi_ls.append(WorkInfo(csvParser.read_csv(work_path + file)))
    result_dict = dict()
    for wi in wi_ls:
        for name, status in wi.data_dict.items():
            if status:
                if name in result_dict.keys():
                    result_dict[name] += 1
                else:
                    result_dict[name] = 1
            else:
                if name not in result_dict.keys():
                    result_dict[name] = 0
    work_count = len(wi_ls)
    for k, v in result_dict.items():
        result_dict[k] = f"{v * 100 / work_count:.2f}%"
    return result_dict


def getExamInfo():
    ei_ls = []
    for file in os.listdir(exam_path):
        if file.endswith(".xlsx") or file.endswith(".xls"):
            ei_ls.append(ExamInfo(excelParser.read_excel(exam_path + file)))
        elif file.endswith(".csv"):
            ei_ls.append(ExamInfo(csvParser.read_csv(exam_path + file)))
    result_dict = dict()
    for ei in ei_ls:
        for name, score in ei.data_dict.items():
            if name in result_dict.keys():
                if score > result_dict[name]:
                    result_dict[name] = score
            else:
                result_dict[name] = score
    return result_dict


def modifyExcel():
    work_result = getWorkCommit()
    exam_result = getExamInfo()
    filename = [file for file in os.listdir(".") if file.endswith(".xlsx") or file.endswith(".xls")]
    if len(filename) == 1:
        excel = Student(excelParser.read_excel(filename[0]))
        for name, commit in work_result.items():
            excel.setCommit(name, commit)
        for name, score in exam_result.items():
            excel.setScore(name, score)
        excelParser.write_excel(output + os.path.basename(filename[0]), excel.getResult())
    else:
        print("待修改文件只能有一个")
        print(ENTER_CONTINUE)
        enterContinue()


def enterContinue():
    result = input()
    while result != "":
        result = input(ENTER_CONTINUE)


def delete():
    delete_set = set()
    for file in os.listdir(delete_path):
        if file.endswith(".csv"):
            di = DeleteInfo(csvParser.read_csv(delete_path + file))
            s = set(di.name_list())
            delete_set = delete_set.union(s)
        elif file.endswith(".xlsx") or file.endswith(".xls"):
            di = DeleteInfo(excelParser.read_excel(delete_path + file))
            s = set(di.name_list())
            delete_set = delete_set.union(s)
    filename = [file for file in os.listdir(".") if file.endswith(".xlsx") or file.endswith(".xls")]
    if len(filename) == 1:
        excel = Student(excelParser.read_excel(filename[0]))
        for item in delete_set:
            excel.data_dict.pop(item)
        excelParser.write_excel(output + os.path.basename(filename[0]), excel.getResult())
    else:
        print("待修改文件只能有一个")
        print(ENTER_CONTINUE)
        enterContinue()


def main_func(num):
    if num == "1":
        print(MENU_MODIFY)
        enterContinue()
        modifyExcel()
    elif num == "2":
        print(MENU_DELETE)
        enterContinue()
        delete()
    elif num == "3":
        png_ls = [file for file in os.listdir(png_path) if
                  file.endswith(".png") or file.endswith(".jpg") or file.endswith(".jpeg")]
        for png in png_ls:
            file_path = pic2excel.image2excel(png)
            file_path = shutil.move(file_path, output + os.path.basename(file_path))
            print("excel文件已生成:" + file_path)
    elif num == "0":
        sys.exit()


def makeDir(*args):
    for item in args:
        if not os.path.exists(item):
            os.makedirs(item)


if __name__ == '__main__':
    err_info = None

    makeDir("结果", "图片", "考试", "作业", "删除")
    while True:
        try:
            print(MENU_MAIN)
            if err_info is not None:
                print(err_info)
                err_info = None
            in_s: str = input("请输入:").strip()
            if in_s.isdigit():
                main_func(in_s)
                print("按enter键继续")
                enterContinue()
            else:
                print("未找到对应项")
        except Exception as e:
            if isinstance(e, CsvReadException):
                print("按enter键继续")
                enterContinue()
            else:
                tb_list = traceback.extract_tb(sys.exc_info()[2])
                err_info = f"刚刚发生了异常:{e.__str__()}\r\n在{os.path.basename(tb_list[0].filename)} {tb_list[0].lineno}行."

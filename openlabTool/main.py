import os

from openlabTool import excelParser, csvParser
from openlabTool.examinfo import ExamInfo
from openlabTool.student import Student
from openlabTool.workinfo import WorkInfo

work_path = "作业/"
exam_path = "考试/"
output = "out/"


def getWorkCommit():
    wi_ls = []
    for file in os.listdir(work_path):
        if file.endswith(".xlsx"):
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
        if file.endswith(".xlsx"):
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


if __name__ == '__main__':
    work_result = getWorkCommit()
    exam_result = getExamInfo()
    filename = [file for file in os.listdir(".") if file.endswith(".xlsx")]
    if len(filename) == 1:
        excel = Student(excelParser.read_excel(filename[0]))
        for name, commit in work_result.items():
            excel.setCommit(name, commit)
        for name, score in exam_result.items():
            excel.setScore(name, score)
        excelParser.write_excel(output + os.path.basename(filename[0]), excel.getResult())
        input("数据添加成功，按任意键结束")
    else:
        print("待修改文件只能有一个")
        input("按任意键结束")

from openlabTool import excelParser
from openlabTool.baseSheet import BaseSheet


class Student(BaseSheet):
    def __init__(self, form: list):
        super().__init__(form)
        score_list = ["得分", "成绩", "分数", "考试", "考试成绩", "考试得分", "考试分数"]
        commit_list = ["作业率", "提交率", "作业提交率"]
        self.__score_index = self.__find_index(score_list, "得分")
        self.__commit_index = self.__find_index(commit_list, "作业率")
        for item in self.data:
            name = item[self.name_index]
            if name not in self.data_dict.keys():
                self.data_dict[name] = item[2:]
            else:
                raise Exception("表单中有重名的同学，请做处理")

    def __find_index(self, keywords, default):
        for item in keywords:
            if item in self.titles:
                return self.titles.index(item)
        self.titles.append(default)
        return len(self.titles) - 1

    def setScore(self, name, score):
        if name in self.data_dict.keys():
            self.data_dict[name][self.__score_index - 2] = score

    def setCommit(self, name, commit):
        if name in self.data_dict.keys():
            self.data_dict[name][self.__commit_index - 2] = commit

    def getResult(self):
        result = []
        for row in self.data:
            name = row[self.name_index]
            ls = [row[0], name]
            ls.extend(self.data_dict[name])
            result.append(ls)
        result.insert(0, self.titles)
        return result

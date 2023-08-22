from openlabTool import excelParser
from openlabTool.baseSheet import BaseSheet


class ExamInfo(BaseSheet):
    def __init__(self, form: list):
        super().__init__(form)
        self.__score_index = self.titles.index("得分")
        for item in self.data:
            name = item[self.name_index]
            if name in self.data_dict.keys():
                raise Exception("重名了，处理去")
            else:
                self.data_dict[name] = item[self.__score_index]

    def getScore(self, name):
        if name in self.data_dict.keys():
            return self.data_dict[name]
        else:
            return None

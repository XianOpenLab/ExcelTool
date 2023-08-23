from openlabTool.baseSheet import BaseSheet


class WorkInfo(BaseSheet):
    def __init__(self, form: list):
        super().__init__(form)
        self.__status_index = self.titles.index("提交状态")
        for item in self.data:
            name = self.name(item)
            if name not in self.data_dict:
                self.data_dict[name] = True if item[self.__status_index] == "已提交" else False
            else:
                raise Exception("作业中有重名的同学，请处理")

    def getStatus(self, name):
        if name in self.data_dict.keys():
            return self.data_dict[name]
        else:
            return None

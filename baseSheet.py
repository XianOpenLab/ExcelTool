class BaseSheet:
    def __init__(self, form: list):
        self.titles = form[0] if form else []
        self.data = form[1:]
        if "姓名" in self.titles:
            self.name_index = self.titles.index("姓名")
        elif "真实姓名" in self.titles:
            self.name_index = self.titles.index("真实姓名")
        else:
            raise Exception("没有姓名字段")
        self.data_dict = dict()

    def name(self, item):
        return item[self.name_index]

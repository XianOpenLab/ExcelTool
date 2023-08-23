from baseSheet import BaseSheet


class DeleteInfo(BaseSheet):
    def __init__(self, form: list):
        super().__init__(form)

    def name_list(self):
        ls = []
        for row in self.data:
            ls.append(self.name(row))
        return ls

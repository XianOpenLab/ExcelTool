class Sheet:
    def __init__(self, filename, form=None):
        if form is None:
            form = []
        self.title: list = form[0]
        self.data: list = form[1:]
        self.filename = filename

    def get_form_data(self):
        row_data = [self.title]
        row_data += self.data
        return row_data

    # def clone(self):
    #     return DataInstance(copy.deepcopy(self.get_row_data()))

    def excel_init(self, name):
        if self.title.__contains__(name):
            for i in range(len(self.data)):
                self.set(i, name, 0)
        else:
            self.title.append(name)
            for row in self.data:
                row.append(None)

    def get(self, index, name):
        if len(self.data) > index and self.title.__contains__(name):
            position = self.title.index(name)
            return self.data[index][position]

    def getCol(self, name):
        if name in self.title:
            index = self.title.index(name)
            return [item[index] for item in self.data]

    def getRow(self, index):
        if 0 <= index < len(self.data):
            return self.data[index]

    def set(self, index, name, value):
        if len(self.data) > index:
            if self.title.__contains__(name):
                position = self.title.index(name)
                self.data[index][position] = value
            else:
                self.title.append(name)
                for row in self.data:
                    row.append(None)
                self.data[index][len(self.title) - 1] = value

    def remove(self, index):
        del self.data[index]

    def title_contain(self, name):
        for key in self.title:
            if key.__contains__(name):
                return True
        return False

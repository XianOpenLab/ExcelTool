class DataInstance:
    def __init__(self, filename, form=None):
        if form is None:
            form = []
        self.title = form[0]
        self.data = form[1:]
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

    def set(self, index, name, value):
        if len(self.data) > index:
            if self.title.__contains__(name):
                position = self.title.index(name)
                self.data[index][position] = value
            else:
                self.title.append(name)
                for row in self.data:
                    row.append(None)
                self.data[index][len(row) - 1] = value

    def remove(self, index):
        del self.data[index]

    def title_contain(self, name):
        for key in self.title:
            if key.__contains__(name):
                return True
        return False

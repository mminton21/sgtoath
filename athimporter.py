import openpyxl

class Athimporter():
    """Combining three reports into a class so that it is easier to execute. Takes multiple columns of data in a row and 
    concatenates to one line separated by '^'."""

    cell_list = []
    cell_dict = []
    new_cell_dict = []
    i = 0

    def __init__(self, filename, worksheet_name, cell_start, cell_end, start_row, column_number):
        self.cell_start = cell_start
        self.cell_end = cell_end
        self.start_row = start_row
        self.column_number = column_number
        self.filename = filename
        self.worksheet_name = worksheet_name
        self.ath = openpyxl.load_workbook(self.filename)
        self.sheet = self.ath[self.worksheet_name]
        self.cell_list = []
        self.cell_dict = []
        self.new_cell_dict = []

    def cell_pull(self):
        for cell_values in self.sheet[self.cell_start:self.cell_end]:
            for co in cell_values:
                if co.value:
                    self.cell_list.append(str(co.value))
            self.cell_dict.append(self.cell_list)
            self.cell_list = []

    def joiner(self):
        self.cell_pull()
        for contents in self.cell_dict:
            separator = "^"
            filler = separator.join(contents)
            self.new_cell_dict.append(filler)

    def exceller(self):
        self.joiner()
        i = 0
        counter = len(self.new_cell_dict)
        while i < counter:
            str_of_list = str(self.new_cell_dict[i])
            self.sheet.cell(row=int(self.start_row) + 1, column=int(self.column_number)).value = str_of_list
            i += 1
            self.start_row += 1

    def finisher(self):
        self.exceller()
        self.ath.save(self.filename)

#tester = Athimporter("test.xlsx","Sheet1","A2","C4",1,5)
#tester.finisher()

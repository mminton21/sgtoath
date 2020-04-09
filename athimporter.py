import openpyxl

class Athimporter():
    """Combining three reports into a class so that it is easier to execute. Takes multiple columns of data in a row and 
    concatenates to one line separated by '^'."""

    #initalizes variables needed, including name of file, name of sheet, start and end cell, start column and row.
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

    #iterates through the data from the cells specified in init and places them in a list.
    def cell_pull(self):
        for cell_values in self.sheet[self.cell_start:self.cell_end]:
            for co in cell_values:
                if co.value:
                    self.cell_list.append(str(co.value))
            self.cell_dict.append(self.cell_list)
            self.cell_list = []

    #iterates through list created in cell_pull and joins each row as strings separated by a "^"
    def joiner(self):
        self.cell_pull()
        for contents in self.cell_dict:
            separator = "^"
            filler = separator.join(contents)
            self.new_cell_dict.append(filler)

    #takes results from joiner and goes through each row and moves them to the column and row specified in sheet.
    def exceller(self):
        self.joiner()
        i = 0
        counter = len(self.new_cell_dict)
        while i < counter:
            str_of_list = str(self.new_cell_dict[i])
            self.sheet.cell(row=int(self.start_row) + 1, column=int(self.column_number)).value = str_of_list
            i += 1
            self.start_row += 1

    #takes work done in exceller and then saves the file.
    def finisher(self):
        self.exceller()
        self.ath.save(self.filename)

import xlrd


class excel:
    'Excel file loding python'
    
    def __init__(self, file):
        self.file = file
        self.workBook = xlrd.open_workbook(file)

    def sheets_name(self):
        return self.workBook.sheet_names()

    def sheet_byname(self, sheet_name):
        return self.workBook.sheet_by_name(sheet_name)
    
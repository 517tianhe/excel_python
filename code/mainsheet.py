import xlrd


class mainsheet:

    @classmethod
    def analysis_sheet(cls, sheet):
        nrows = sheet.nrows
        ncols = sheet.ncols
        print('sheet的名字以及行数和列数:', sheet.name, ', ', nrows, ', ', ncols)
        for row in range(0, nrows):
            print(sheet.row_values(row))

from excel import excel
from mainsheet import mainsheet


def read_excel(file):
    workbook = excel(file)
    print('打印所有的sheet表名')
    names = workbook.sheets_name()
    print(names)
    main = names[0]
    if main == '总表':
        mainsheet.analysis_sheet(workbook.sheet_byname(main))
    else:
        print('没有总表sheet')


if __name__ == '__main__':
    file = 'example_excel.xlsx'
    read_excel(file)

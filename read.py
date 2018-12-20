import xlrd
from datetime import date 


def read_excel(file):

    # 打开文件
    workbook = xlrd.open_workbook(file)

    # 打印所有的sheet表名
    print(workbook.sheet_names())
    print(workbook.sheets())

    # 得到第一张表的名字
    sheet1_name = workbook.sheet_names()[0]

    # 根据sheet索引或者名称获取sheet的内容
    sheet1_by_list = workbook.sheets()[0]
    sheet1_by_index = workbook.sheet_by_index(0)  # sheet索引从0开始
    sheet1_by_name = workbook.sheet_by_name(sheet1_name)

    # 获取sheet的名称和行、列数
    print(sheet1_by_list.name, sheet1_by_list.nrows, sheet1_by_list.ncols, \
          sheet1_by_index.name, sheet1_by_index.nrows, sheet1_by_index.ncols,\
          sheet1_by_name.name, sheet1_by_name.nrows, sheet1_by_name.ncols)

    # 获取整行或者整列的值
    rows = sheet1_by_name.row_values(0)  # 获取第一行的值
    cols = sheet1_by_name.col_values(0)  # 获取第一列的值
    print(rows, cols)

    # 获取单元格内容,A\b\c\D在excel中是第1、2、3、4列
    cell_A1 = sheet1_by_name.cell(0, 0).value
    cell_C1 = sheet1_by_name.cell(0, 2).value
    cell_B1 = sheet1_by_name.row(0)[1].value
    cell_D2 = sheet1_by_name.col(3)[1].value
    print(cell_A1, cell_C1, cell_B1, cell_D2)

    # 获取单元格的数据类型
    # ctype:0:empty;1:string;2:number; 3:date;4:boolean;5:error 
    print('cell(0,0)数据类型:', sheet1_by_name.cell(0, 0).ctype)
    print('cell(1,0)数据类型:', sheet1_by_name.cell(1, 0).ctype)
    print('cell(1,1)数据类型:', sheet1_by_name.cell(1, 1).ctype) 
    print('cell(1,2)数据类型:', sheet1_by_name.cell(1, 2).ctype)
    # 如果返回的值是0，说明这个单元格的值是空值

    # 获取单元格内容为日期的数据
    date_value = xlrd.xldate_as_tuple(sheet1_by_name.cell_value(2, 2), workbook.datemode)
    print(date_value)
    # 生成日期
    print(date(*date_value[:3]))
    # 转化为标准格式
    # 首先对单元格的内容做一个判断处理，然后做时间修改
    if (sheet1_by_name.cell(2, 2).ctype == 3):  
        date_value = xlrd.xldate_as_tuple(sheet1_by_name.cell_value(2, 2), workbook.datemode) 
        print(date(*date_value[:3]).strftime('%Y/%m/%d'))

    # 对合并单元格的处理，读取合并单元格值
    # 合并单元格起始项是可以索引到值 
    print('cell(1,4)值:', sheet1_by_name.cell(1, 4).value)  # 朋友
    # 合并单元格起始项是不可以索引到值 
    print('cell(2,4)值:', sheet1_by_name.row(2)[4].value)     # 空值
    print('cell(2,4)数据类型:', sheet1_by_name.cell(2, 4).ctype)

    # 获取合并的单元格
    '''
    读取整个excel文件的时候需要将formatting_info参数设置为True，默认是False；
    workbook = xlrd.open_workbook(file,formatting_info=True) 
    所以上面获取合并的单元格数组为空：print(sheet1_by_name.merged_cells)   #返回[]空列表
    '''
    print(sheet1_by_name.merged_cells)  
    '''merged_cells返回的这四个参数的含义是：(row,row_range,col,col_range),
    其中[row,row_range)包括row,不包括row_range;col也是一样
    (1, 3, 4, 5)的含义是：第1到2行（不包括3）合并，(7, 8, 2, 5)的含义是：第2到4列合并。
    '''
    # 利用上述原理，可以分别获取合并的二个单元格的内容：
    merged_cells = sheet1_by_name.merged_cells
    for cell in merged_cells:
        print('cell合并值:', sheet1_by_name.cell(cell[0], cell[2]).value)


if __name__ == '__main__':
    file = 'example_excel.xlsx'
    read_excel(file)

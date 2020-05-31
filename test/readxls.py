import xlrd #xlrd是读excel
import xlwt #xlwt是写excel的库

def readxls():
    workbook = xlrd.open_workbook(r'books.xls',formatting_info=True) #要获得合并单元格，读取文件的时候需要将formatting_info参数设置为True，默认是False
    print(workbook.sheet_names()) #输出标签名称
    #sheet2 = workbook.sheet_by_index(0) #打开第一个标签
    sheet2 = workbook.sheet_by_name('A')  # 打开标签'A'
    nrows = sheet2.nrows #获取行数
    ncols = sheet2.ncols #获取列数
    print(nrows, ncols) #输出结果

    # 获取整行和整列的值（数组）
    rows = sheet2.row_values(2)  # 获取第三行内容
    cols = sheet2.col_values(1)  # 获取第二列内容
    print(rows)
    print(cols)

    #ctype : 0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
    cell_A = sheet2.cell(1,1).value #取出第二行第二列值
    print(sheet2.cell(1,1).ctype,cell_A) #输出这个值

    cell_B = sheet2.cell(1,2).value #取出第二行第三列值
    print(sheet2.cell(1,2).ctype,cell_B) #输出这个值

    cell_C = sheet2.cell(2,1).value #取出第三行第二列值
    print(sheet2.cell(2,1).ctype,cell_C) #输出这个值

    #找到合并的单元格并打印
    print(sheet2.merged_cells)

    #获取merge_cells返回的row和col低位的索引即可
    for (rlow,rhigh,clow,chigh) in sheet2.merged_cells:
        print(sheet2.cell(rlow,clow).value)

if __name__ == '__main__':
    readxls()
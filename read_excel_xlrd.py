# coding:utf-8
import xlrd


def to_int(n):
    try:
        if n is None:
            return None
        if isinstance(n, int):
            return n
        if isinstance(n, float):
            return int(n)
        if isinstance(n, str):
            n = n.replace('\'', '')
            return int(float(n))
        if isinstance(n, unicode):
            n = n.replace('\'', '')
            return int(float(n))

        return int(float(str(n)))
    except ValueError as e:
        print("to_int Error:" + str(e))
    return None


def read_excel_xlrd(f):
    excel = xlrd.open_workbook(f)
    table = excel.sheets()[0]       # 获取工作簿sheet列表的第一项（第一个标签页
    row = table.nrows               # 获取最大行数
    col = table.ncols               # 获取最大列数

    # 如果行数或者列数为0则判定为空工作簿，返回
    if row == 0 or col == 0:
        return 0, {}

    # 从最大行的第一列循环往回检查每一行，如果有空，则返回
    while table.cell(0, col-1).value == u"" or table.cell(0, col-1).value is None:
        # 忽略掉第一列（第一列都是列名
        col -= 1
        if col == 0:
            return None

    # 如果非空，获取内容并返回
    result = {}
    for r in range(1, row):
        local_row = []
        for c in range(0, col):
            # 从第二行第一列开始读取到第一行最后一列
            # ↓ctype：获取单元格的数据类型，XL_CELL_NUMBER：数字类型，XL_CELL_DATE：日期类型，XL_CELL_TEXT：文本类型
            if table.cell(r, c).ctype == xlrd.XL_CELL_NUMBER and int(table.cell(r, c).value) == table.cell(r, c).value:
                local_row.append(table.cell(r, c).value)    # 如果为数字类型且为整数，不做处理直接添加到local_row列表中
            elif xlrd.XL_CELL_DATE == table.cell(r, c).ctype:
                showval = xlrd.xldate_as_tuple(table.cell(r, c).value, excel.datemode)
                # 若年月日都为零，以格式(12:00:00)处理, 否则以日期格式(2016-12-12)处理
                if showval[0] == 0 and showval[1] == 0 and showval[2] == 0:
                    cell_value = "\'%d:%02d:%02d" % (showval[3], showval[4], showval[5])
                    cell_value = cell_value.decode('utf-8')
                    local_row.append(cell_value)
                else:
                    cell_value = "%4d-%02d-%02d" % (showval[0], showval[1], showval[2])
                    cell_value = cell_value.decode('utf-8')
                    local_row.append(cell_value)
            elif xlrd.XL_CELL_TEXT == table.cell(r, c).ctype:
                # 若为文本单元格，保存格式为： '文本 ，（前边加上
                if table.cell(r, c).value == u"" or table.cell(r, c).value == "" or table.cell(r, c).value is None:
                    local_row.append(None)
                else:
                    local_row.append("\'" + table.cell(r, c).value)
            elif xlrd.XL_CELL_EMPTY == table.cell(r, c).ctype or xlrd.XL_CELL_BLANK == table.cell(r, c).ctype:
                # 若为空单元格，row_local添加None
                local_row.append(None)
            else:
                if table.cell(r, c).value == u"" or table.cell(r, c).value == "" or table.cell(r, c).value is None:
                    local_row.append(None)
                else:
                    local_row.append(str(table.cell(r, c).value))

        rowid = to_int(local_row[0])
        if rowid is not None:
            local_row[0] = rowid
            result[rowid] = local_row

    return col, result


print(read_excel_xlrd(r'C:\Users\Administrator\Desktop\ExcelMerger\test.xlsx'))


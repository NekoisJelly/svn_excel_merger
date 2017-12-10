# coding:utf-8
import win32com.client as win32


def read_excel_win32com(f):
    # 参考文档见https://msdn.microsoft.com/vba/vba-excel
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = 0
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(f)
    ws = wb.Worksheets(1)
    try:
        # range对象：工作表中的一个单元格或者一片区域
        # 大部分操作都是对于range对象的操作
        # 参考：https://msdn.microsoft.com/vba/vba-excel/articles/range-style-property-excel
        used = ws.UsedRange
        row = used.Rows.Count
        col = used.Columns.Count

        # if this is an empty excel
        if row == 0 or col == 0:
            return 0, {}

        while ws.Cells(1, col-1).Value == u'' or ws.Cells(1, col-1).Value is None:
            col -= 1
            if col == 0:
                return None

        # result = {}
        # for r in range(1, row):
        #     local_row = []
        #     for c in range(0, col):
        #         # if table.cell(r, c).ctype ==
        print(ws.Cells(1, col).Value)
        print(ws.Cells(1, col).getCellType)

        wb.Close()
        return row, col

    except Exception as e:
        print(e)
        print("[修改Excel文件失败]:" + f)


print(read_excel_win32com(r'C:\Users\Administrator\Desktop\ExcelMerger\test.xlsx'))
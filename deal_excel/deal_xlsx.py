"""
使用openpyxl模块处理xlsx文件
可直接使用openpyxl操作xlsx文件，不必像操作xls文件一样需要xlutils模块作为中间转化
"""
import datetime
from copy import copy

import openpyxl


def read_xlsx(filepath):
    """
    使用openpyxl读取xlsx文件内容
    """
    # 获取工作簿
    wb = openpyxl.load_workbook(filepath)
    # 获取默认sheet
    ws = wb.active
    # 或者根据sheet名字获取sheet
    # ws = wb.get_sheet_by_name(sheet_name)
    # ws = wb.[sheet_name]

    # ws.rows或ws.columns获取表格的行或列的生成器
    for row in ws.rows:
        for cell in row:
            print(cell.value)

    # 获取某一单元格
    cell = ws["A1"]
    cell = ws.cell(1, 1)
    print(cell.value)

    # 获取多单元格
    # 1. 切片
    cells = ws["A1":"C2"]  # 获得每行单元格元祖组成的元祖,如：((1行1列，1行2列，1行3列),(2行1列，2行2列，2行3列))
    for row in cells:
        for cell in row:
            print(cell.value)
    # 2.指定行或列
    row1 = ws[1]  # col1 = ws["A:C"]，结果元祖以行为单位
    for cell in row1:
        print(cell.value)
    row1to2 = ws[1:2]  # col1to3 = ws["A:C"]，注意结果元祖以列为单位
    for row in row1to2:
        for cell in row:
            print(cell.value)


def create_xlsx(filepath):
    """
    使用openpyxl新建xlsx文件
    """
    # 创建工作簿
    wb = openpyxl.Workbook()
    ws = wb.create_sheet("sheet1", 0)  # 创建sheet，可指定位置
    ws2 = wb.create_sheet("sheet2", 1)

    datas = [
        ("标题1", "标题2", "标题3", "标题4"),
        ("内容11", "内容12", "内容13", "内容14"),
        ("内容21", "内容32", "内容23", "内容24"),
    ]

    # 插入行
    for row_data in datas:
        ws.append(row_data)

    # 直接修改单元格值
    ws.cell(2, 1, 'sss')

    wb.save(filepath)


def modify_xlsx(filepath):
    """
    使用openpyxl修改xlsx文件
    """
    wb = openpyxl.load_workbook(filepath)
    ws = wb["Sheet1"]
    new_sheet = wb.create_sheet("sheet_create")
    # wb.remove(new_sheet)  # 删除sheet
    for i in range(5):
        new_sheet.append(range(4))
    # 合并单元格,仅保留左上角单元格数据
    new_sheet.merge_cells("A1:B2")
    # 拆分单元格,数据填充在左上角单元格，其它单元格无数据
    new_sheet.unmerge_cells("A1:B2")

    nrows = ws.max_row  # 获取行数
    ncols = ws.max_column  # 获取列数
    for row in range(0, nrows):
        for col in range(0, ncols):
            cell = ws.cell(row + 1, col + 1)
            cell_value = cell.value
            if isinstance(cell_value, datetime.datetime):
                cell.value = cell_value + datetime.timedelta(days=1)

    wb.save(filepath)


def copy_temp_xlsx(filepath):
    """
    使用openpyxl复制单元格格式
    """
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active
    new_ws1 = wb.copy_worksheet(ws)  # 拷贝sheet，部分格式肯能无法拷贝
    new_ws1.title = "直接拷贝"

    wb.save(filepath)


if __name__ == '__main__':
    filepath = "excel_ouput/temp_xlsx.xlsx"
    new_path = "excel_ouput/new_xlsx.xlsx"
    # read_xlsx(filepath)
    # create_xlsx(new_path)
    # modify_xlsx(filepath)
    copy_temp_xlsx(filepath)

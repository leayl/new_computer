"""
python新建、读取、修改xls文件方法
1.新建
  新建xls文件需要使用xlwt模块，该模块仅可新建写入文件，不可读取。以下new_xls(filename)函数。
2.读取
  读取xls文件需要使用xlrd模块，该模块仅可读取文件内容，不可修改写入。以下read_xls(filename)函数。
3.修改
  由于xlwt、xlrd模块只能新建或读取xls文件，需要使用xlutils模块将xlrd读取的内容
  复制并转化为xlwt对象，对其进行修改后保存，保存为原名则为修改，否则为另存为。
  以下modify_xls(filename)函数。
  如：
  需要使用某xls文件作为模板新建文件时，需要先使用xlrd读取出所需要的格式，
  再在使用xlutils转化的xlwt对象写入文件时使用该格式。
"""
import os
import xlwt
import xlrd
from xlutils.copy import copy
from xlutils.filter import process, XLWTWriter, XLRDReader
import datetime


def new_xls(filename):
    """
    新创建xls文件
    :param filename:
    :return:
    """
    wb = xlwt.Workbook()
    ws = wb.add_sheet('sheet')
    for row in range(5):
        for col in range(5):
            ws.write(row, col, f'行{row + 1}列{col + 1}')
    wb.save(filename)


def read_xls(filename):
    """
    读取xls文件
    :param filename:
    :return:
    """
    rb = xlrd.open_workbook(filename, formatting_info=True)
    rs = rb.sheet_by_index(0)
    nrows = rs.nrows
    ncols = rs.ncols
    for row in range(nrows):
        print()
        for col in range(ncols):
            value = rs.cell_value(row, col)  # 根据行列索引获取指定单元格的值，或value = re.cell(row, col).value
            print(value, end=' ')


def modify_xls(filename):
    """
    修改已存在的xls文件
    :param filename:
    :return:
    """
    workbook = xlrd.open_workbook(filename, formatting_info=True)  # 首先使用xlrd读取文件，只可读，不可写
    sheet = workbook.sheet_by_index(0)
    nrows = sheet.nrows
    ncols = sheet.ncols
    rowdatas = []
    for row in range(nrows):
        rowdata = sheet.row_values(row)
        rowdatas.append(rowdata)

    new_workbook = copy(workbook)  # 使用xlutils拷贝原工作簿，将xlrd对象转化为xlwt对象，只可写，不可读
    new_sheet = new_workbook.get_sheet(0)  # 使用get_sheet方法获取新的工作簿的sheet
    title_style, content_style = define_style()
    for row_index in range(nrows):
        if row_index == 0:
            style = title_style
        else:
            style = content_style
        for col_index in range(ncols):
            new_sheet.write(row_index, col_index, rowdatas[row_index][col_index], style)
    output_file = os.path.join(os.path.dirname(filename), 'define_style.xls')
    new_workbook.save(output_file)  # 保存为原文件名则覆盖原文件，否则另存为其他文件


def define_style():
    """
    返回自定义表格标题和内容的格式
    :param filename:
    :return:
    """
    # 背景色
    title_pattern = xlwt.Pattern()
    title_pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # 有背景色
    title_pattern.pattern_fore_colour = 3  # 设置背景颜色

    # 字体
    title_font = xlwt.Font()
    title_font.bold = True
    title_font.shadow = True
    title_font.height = 280  # 200为10号字体
    title_font.name = "Times New Roman"
    title_font.colour_index = 5  # 字体颜色 0：黑，1：白，2：红，3：绿，4：蓝，5：黄

    # 对齐
    title_alignment = xlwt.Alignment()
    title_alignment.vert = xlwt.Alignment.VERT_CENTER
    title_alignment.horz = xlwt.Alignment.HORZ_CENTER

    content_pattern = xlwt.Pattern()
    content_pattern.pattern_fore_colour = 3

    content_font = xlwt.Font()
    content_font.height = 200  # 200为10号字体
    content_font.family = xlwt.Font.FAMILY_SWISS

    content_alignment = xlwt.Alignment()
    content_alignment.vert = xlwt.Alignment.VERT_BOTTOM
    content_alignment.horz = xlwt.Alignment.HORZ_LEFT

    # 边框
    border = xlwt.Borders()
    border.left = xlwt.Borders.THIN_DASH_DOT_DOTTED
    border.right = xlwt.Borders.THIN_DASH_DOT_DOTTED
    border.top = xlwt.Borders.THIN_DASH_DOT_DOTTED
    border.bottom = xlwt.Borders.THIN_DASH_DOT_DOTTED

    title_style = xlwt.XFStyle()
    title_style.pattern = title_pattern
    title_style.font = title_font
    title_style.borders = border
    title_style.alignment = title_alignment

    content_style = xlwt.XFStyle()
    content_style.pattern = content_pattern
    content_style.font = content_font
    content_style.borders = border
    content_style.alignment = content_alignment

    return title_style, content_style


def new_xls_with_template_style(filename):
    """
    使用某xls文件内的格式创建新的单元格
    :param filename:
    :return:
    """
    rb = xlrd.open_workbook(filename, formatting_info=True)
    rs = rb.sheet_by_index(0)

    # 获取要从模板文件中获得的样式
    writer = XLWTWriter()
    process(XLRDReader(rb, 'nothing.xls'), writer)  # 'noting.xls'为任意命名，实际不会产生影响
    title_style = writer.style_list[rs.cell_xf_index(0, 0)]  # 第一行一列单元格格式
    content_style = writer.style_list[rs.cell_xf_index(1, 0)]  # 第二行一列单元格格式
    content_style.num_format_str = 'YYYY/MM/DD'  # 防止日期格式丢失，特意设置格式

    new_b = copy(rb)
    new_s = new_b.get_sheet(0)

    nrows = 10
    ncols = 12
    for row_index in range(nrows):
        style = content_style
        cell_index = datetime.date.today()
        if row_index == 0:
            style = title_style
            cell_index = '标题'
        for col_index in range(ncols):
            if isinstance(cell_index, str):
                new_s.write(row_index, col_index, cell_index, style)
            else:
                new_s.write(row_index, col_index, cell_index + datetime.timedelta(days=1), style)
    new_file = os.path.join(os.path.dirname(filename), 'from_template_style.xls')
    new_b.save(new_file)


if __name__ == '__main__':
    file = "temp_xls.xls"
    file_dir = os.path.join(os.path.dirname(__file__), 'excel_ouput')
    if not os.path.exists(file_dir):
        os.mkdir(file_dir)
    filename = os.path.join(file_dir, file).replace('\\', '/')
    # new_xls(filename)
    # read_xls(filename)
    # modify_xls(filename)
    new_xls_with_template_style(filename)

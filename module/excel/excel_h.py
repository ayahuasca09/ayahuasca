import config
import openpyxl
from pprint import pprint
import os
import module.oi.oi_h as oi_h

"""打开工作簿并获取工作表"""


def excel_get_sheet(path, sheetname):
    # 打开工作簿
    wb = openpyxl.load_workbook(path)
    # print(wb)
    # 获取工作表
    sheet = wb[sheetname]
    # print(sheet)
    return sheet, wb


"""获取表头数据所在列"""


def excel_get_sheet_title_column(sheet, title_list):
    # 标题所在列的字典映射
    title_colunmn_dict = {}
    # 遍历第一行
    for cell in list(sheet.rows)[0]:
        for title_name in title_list:
            if title_name in str(cell.value):
                title_colunmn_dict[title_name] = cell.column
    return title_colunmn_dict


"""获取当前目录下要遍历的表路径列表"""


def excel_get_path_list(py_path):
    file_names = []
    for i in os.walk(py_path):
        file_names.append(i)
    # pprint("输出文件夹下的文件名：")
    file_name_list = file_names[0][2]
    return file_name_list


"""获取标题所在列"""


def sheet_title_column(title_name, title_colunmn_dict):
    file_name = ""
    title_colunmn = ""
    if title_name in title_colunmn_dict.keys():
        # 读取测试
        title_colunmn = title_colunmn_dict[title_name]
    else:
        oi_h.print_error(title_name + ":标题列不存在，请检查表格中是否有该标题的列表")

    return title_colunmn

import config
import openpyxl
from pprint import pprint

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

import openpyxl
import json
import sys
from os.path import abspath, dirname
import os
from pprint import pprint
import re

# 文件所在目录
py_path = ""
if hasattr(sys, 'frozen'):
    py_path = dirname(sys.executable)
elif __file__:
    py_path = dirname(abspath(__file__))


# 检测字符串是否含有中文
def check_is_chinese(name):
    for ch in name:
        if u'\u4e00' <= ch <= u'\u9fff':
            return True
    return False


"""获取.xlsx文件"""
file_names = []
for i in os.walk(py_path):
    file_names.append(i)
# pprint("输出文件夹下的文件名：")
file_name_list = file_names[0][2]

# 提取规则：只提取xlsx文件
for i in file_name_list:
    if ".xlsx" in i:
        # 拼接xlsx的路径
        file_path_xlsx = os.path.join(py_path, i)
        # 获取xlsx的workbook
        wb = openpyxl.load_workbook(file_path_xlsx)
        # 获取xlsx的所有sheet
        sheet_names = wb.sheetnames
        # 加载所有工作表
        for sheet_name in sheet_names:
            sheet = wb[sheet_name]
            # 获取工作表第一行数据
            for cell in list(sheet.rows)[0]:
                if 'Sample Name' in str(cell.value):
                    # 获取音效名下的内容
                    for cell_sound in list(sheet.columns)[cell.column - 1]:
                        # 空格和中文不检测
                        if cell_sound.value != None:
                            if check_is_chinese(cell_sound.value) == False:
                                """❤❤❤❤数据获取❤❤❤❤"""
                                pprint(cell_sound.value)

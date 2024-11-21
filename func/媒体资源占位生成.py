import openpyxl
import json
import sys
from os.path import abspath, dirname
import os
from pprint import pprint
import re
from waapi import WaapiClient
import shutil
from openpyxl.cell import MergedCell
import module.cloudfeishu.cloudfeishu_h as cloudfeishu_h
import module.oi.oi_h as oi_h
import config
import module.excel.excel_h as excel_h
from pathlib import Path
import 命名规范检查

"""根目录获取"""
# 获取当前脚本的文件名及路径
script_name_with_extension = __file__.split('/')[-1]
# 去掉扩展名
script_name = script_name_with_extension.rsplit('.', 1)[0]
# 替换为files
root_path = script_name.replace("func", "files")
# print(f"当前脚本的名字是: {root_path}")c

file_name_list = excel_h.excel_get_path_list(root_path)
# pprint(file_name_list)

"""在线表获取"""
# 规范检查表
# for excel_media_token in config.excel_media_token_list:
#     _, excel_name = cloudfeishu_h.get_excel_token(excel_media_token)
#     # pprint(excel_name)
#     for file_name in file_name_list:
#         if excel_name in file_name:
#             print(excel_name)
#             cloudfeishu_h.download_cloud_sheet(excel_media_token, os.path.join(root_path, file_name))
#             break

"""收集unit的路径"""
audio_unit_list = []
event_unit_list = []

"""*****************功能检测区******************"""
"""获取表格中的事件描述和状态所在的列"""


def get_descrip_and_status_column():
    require_module_column = None
    second_module_column = None
    require_name_column = None
    status_column = None
    sample_name_column = None

    if list(sheet.rows)[0]:
        for cell in list(sheet.rows)[0]:
            if cell.value:
                if 'Require Module' in str(cell.value):
                    require_module_column = cell.column
                    # print(require_module_column)
                elif 'Second Module' in str(cell.value):
                    second_module_column = cell.column
                    # print(second_module_column)
                elif 'Require Name' in str(cell.value):
                    require_name_column = cell.column
                    # print(require_name_column)
                elif 'Status' in str(cell.value):
                    status_column = cell.column
                    # print(status_column)
                elif 'Sample Name' in str(cell.value):
                    sample_name_column = cell.column
                    # print(status_column)

    return require_module_column, second_module_column, require_name_column, status_column, sample_name_column


with WaapiClient() as client:
    """*****************Wwise功能区******************"""

    """*****************主程序处理******************"""

    # 记录所有资源名称
    sound_name_list = []

    # 提取规则：只提取xlsx文件
    for i in file_name_list:
        if ".xlsx" in i:
            # 拼接xlsx的路径
            file_path_xlsx = os.path.join(root_path, i)
            # 获取xlsx的workbook
            wb = openpyxl.load_workbook(file_path_xlsx)
            # 获取xlsx的所有sheet
            sheet_names = wb.sheetnames
            # 加载所有工作表
            for sheet_name in sheet_names:
                sheet = wb[sheet_name]

                require_module_column, second_module_column, require_name_column, status_column, sample_name_column = get_descrip_and_status_column()
                if sample_name_column:
                    for cell_sound in list(sheet.columns)[sample_name_column - 1]:
                        if status_column:
                            if sheet.cell(row=cell_sound.row,
                                          column=status_column).value in config.status_list:
                                # pprint(cell_sound.value)
                                if 命名规范检查.check_basic(cell_sound.value, audio_unit_list, event_unit_list):
                                    # 通过命名规范检查
                                    # pprint(cell_sound.value)
                                    pass

# # unit列表长度排序
audio_unit_list.sort(key=len)
event_unit_list.sort(key=len)
pprint(audio_unit_list)
pprint(event_unit_list)

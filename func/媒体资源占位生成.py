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
import config
import module.excel.excel_h as excel_h
from pathlib import Path

"""根目录获取"""
# 获取当前脚本的文件名及路径
script_name_with_extension = __file__.split('/')[-1]
# 去掉扩展名
script_name = script_name_with_extension.rsplit('.', 1)[0]
# 替换为files
root_path = script_name.replace("func", "files")
# print(f"当前脚本的名字是: {root_path}")

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

                # 获取表头所有数据所在列
                title_colunmn_dict = excel_h.excel_get_all_sheet_title_column(sheet)


                # for title_colunmn in title_colunmn_dict:
                #     if title_colunmn:
                #         print(title_colunmn)
                #         if "Sample Name" in title_colunmn:
                #             for cell in list(sheet.columns)[title_colunmn_dict[title_colunmn]-1]:
                #                 print(1)
                #             break


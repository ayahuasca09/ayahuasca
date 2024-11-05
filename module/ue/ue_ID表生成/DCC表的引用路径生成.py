from waapi import WaapiClient
import os
import sys
import re
import pandas as pd
import tkinter as tk
from tkinter import messagebox
from os.path import abspath, dirname
from pprint import pprint
import json

# 自定义库
import module.excel.excel_h as excel_h
import module.config as config
import module.oi.oi_h as oi_h
import module.waapi.waapi_h as waapi_h
import module.json.json_h as json_h

"""**********************数据获取*************************"""
sheet, wb = excel_h.excel_get_sheet(config.excel_dcc_dt_audio_path, config.dt_audio_sheet_name)

"""*****************功能函数**********************"""
"""查找ID的引用路径"""


def find_ID_refer_path(json_path):
    file_dict, _ = oi_h.get_type_file_name_and_path('.json', json_path)
    # 存储id所在的引用映射路径
    id_refer_dict = {}
    # pprint(file_dict)
    for json_name in file_dict:
        with open(file_dict[json_name], 'r') as jsfile:
            # 这是json转换后的dict，可在python里处理了
            js_dict = json.load(jsfile)
            if "taskConfigData" in js_dict:

                if js_dict["taskConfigData"]:
                    for state in js_dict["taskConfigData"]:
                        state_dict = js_dict["taskConfigData"][state]
                        # pprint(state_dict)
                        if "configList" in state_dict:
                            if state_dict["configList"]:
                                for config_dict in state_dict["configList"]:
                                    # pprint(config_dict)
                                    if "_ClassName" in config_dict:
                                        ue_file_name = re.sub(r'_Skill\.json$', '', json_name)
                                        if config_dict["_ClassName"] == "/Script/SPGameFramework.ANGEConfigPlayAudio":
                                            if (config_dict["eventId"] != 0) and (config_dict["eventId"] != -1):
                                                id_refer_dict[config_dict["eventId"]] = (
                                                    state, ue_file_name)
                                            if "endEventId" in config_dict:
                                                if (config_dict["endEventId"] != 0) and (
                                                        config_dict["endEventId"] != -1):
                                                    id_refer_dict[config_dict["endEventId"]] = (
                                                        state, ue_file_name)
                                                # print(config_dict["endEventId"])

    # pprint(id_refer_dict)
    return id_refer_dict


"""*****************主程序**********************"""
# 查找ID的引用路径
id_refer_dict = find_ID_refer_path(config.ue_ge_path)
# find_ID_refer_path(r"S:\Ver_1.0.0\Project\Content\SPSkill\BossAshley_Attack1")
# {21: ('state1', 'Candleman_01_c_Stun_HitR'),
#  183: ('state1', 'GratiaShieldATK_CharSkin'),
#  195: ('state1', 'CandlemanMeleeCharSkin'),}

# 获取表头所有数据所在列
title_colunmn_dict = excel_h.excel_get_all_sheet_title_column(sheet)
# print(title_colunmn_dict)

# 然后需要把数据写入dcc的表
# 遍历第一列并更新第八列
for row in range(2, sheet.max_row + 1):  # 从第二行开始遍历，假设第一行为标题
    first_col_value = sheet.cell(row=row, column=1).value
    if first_col_value in id_refer_dict:
        if '配置路径' in title_colunmn_dict:
            # 将元组转str
            sheet.cell(row=row, column=title_colunmn_dict['配置路径']).value = ','.join(id_refer_dict[first_col_value])
        if '配置方式' in title_colunmn_dict:
            sheet.cell(row=row, column=title_colunmn_dict['配置方式']).value = 'GE'
wb.save(config.excel_dcc_dt_audio_path)

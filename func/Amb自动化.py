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
import module.waapi.waapi_h as waapi_h
from pathlib import Path

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

"""文件获取"""
# 规范检查表
excel_amb_path = os.path.join(root_path, config.excel_amb_config_path)
# print(wb)

"""在线表获取"""
# 规范检查表
# check_name_token = config.amb_sheet_token
# cloudfeishu_h.download_cloud_sheet(check_name_token, excel_amb_path)

"""功能数据初始化"""
# switch创建列表
switch_create_list = []
# unit创建列表
unit_create_list = []
# rnd创建列表
rnd_create_list = []

"""*****************功能检测区******************"""

with WaapiClient() as client:
    """*****************Wwise功能区******************"""
    """obj创建"""


    def create_wwise_obj(type, parent, name, notes):
        unit_object = None
        if name != "Amb":
            pprint(type)
            if type == "RandomSequenceContainer":
                pprint(name)
                args = waapi_h.args_rnd_create(parent,
                                               type, name, notes)
            else:
                args = waapi_h.args_object_create(parent,
                                                  type, name, notes)
            unit_object = client.call("ak.wwise.core.object.create",
                                      args)
            oi_h.print_warning(name + "：" + type + "已创建")

        return unit_object


    """获取Wwise中的Unit列表"""


    def get_wwise_type_list(root_path, type):
        # 获取Wwise所有的Audio Unit
        wwise_bus_name_all_list = client.call("ak.wwise.core.object.get",
                                              waapi_h.waql_by_type(type, root_path),
                                              options=config.options)['return']
        wwise_bus_name_name_dict = {item['name']: item['id'] for item in wwise_bus_name_all_list}
        # pprint(wwise_bus_name_name_dict)
        # {'Amb': '{4BB49221-4DC1-4A05-A57C-3D42FA570675}',
        #  'Amb_A01': '{415AC456-D390-43D5-A09E-B5D30B253D76}',
        #  'CG': '{BBB332B4-C230-4D60-9ABA-91A36D968C11}'}
        return wwise_bus_name_name_dict


    """amb创建"""


    def create_wwise_amb(rnd_list, switch_list):
        # switch创建
        for amb_name in switch_list:

            switch_have_dict = get_wwise_type_list(config.wwise_amb_state_path, "SwitchContainer")

            if amb_name not in switch_have_dict:
                # 找到的父级为unit
                unit_parent = oi_h.find_longest_prefix_key(amb_name, switch_have_dict)
                if not unit_parent:
                    switch_object = create_wwise_obj("SwitchContainer",
                                                     config.wwise_amb_global_path, amb_name,
                                                     "")
                else:
                    switch_have_dict = get_wwise_type_list(config.wwise_amb_state_path, "SwitchContainer")
                    switch_object = create_wwise_obj("SwitchContainer",
                                                     switch_have_dict[unit_parent], amb_name,
                                                     "")
        # rnd创建
        for amb_name in rnd_list:
            switch_have_dict = get_wwise_type_list(config.wwise_amb_state_path, "SwitchContainer")
            rnd_have_dict = get_wwise_type_list(config.wwise_amb_state_path, "RandomSequenceContainer")
            if amb_name not in rnd_have_dict:
                unit_parent = oi_h.find_longest_prefix_key(amb_name, switch_have_dict)
                if not unit_parent:
                    rnd_object = create_wwise_obj("RandomSequenceContainer",
                                                  config.wwise_amb_global_path, amb_name,
                                                  "")
                else:
                    rnd_object = create_wwise_obj("RandomSequenceContainer",
                                                  switch_have_dict[unit_parent], amb_name,
                                                  "")


    """获取需要创建的amb列表"""


    def get_create_amb_name_list():
        # 遍历每一行
        for row in amb_config_sheet.iter_rows():
            amb_name = None
            # 遍历每一列
            for cell in row:
                cell_value, is_merge = excel_h.check_is_mergecell(cell, amb_config_sheet)
                if cell_value:
                    is_unit = False
                    # print(cell_value)
                    if not amb_name:
                        amb_name = cell_value
                        is_unit = True
                    else:
                        amb_name += "_" + cell_value
                        fill = cell.fill
                        if fill.fill_type == 'solid':
                            if fill.start_color.index == "FFD9F3FD":
                                is_unit = True
                    is_switch = False

                    if is_unit:
                        if amb_name not in unit_create_list:
                            unit_create_list.append(amb_name)
                    # 是否为switch容器检查
                    if is_merge:
                        is_switch = True
                    else:
                        if amb_config_sheet.cell(row=cell.row,
                                                 column=cell.column + 1).value:
                            is_switch = True
                    if is_switch:
                        if amb_name not in switch_create_list:
                            switch_create_list.append(amb_name)
                    else:
                        if amb_name not in rnd_create_list:
                            rnd_create_list.append(amb_name)

            # print()


    """*****************主程序处理******************"""
    """环境音结构创建"""

    # 获取xlsx的workbook
    wb = openpyxl.load_workbook(excel_amb_path)
    amb_config_sheet = wb[config.sheet_amb_config]
    get_create_amb_name_list()
    switch_create_list.sort(key=len)
    rnd_create_list.sort(key=len)
    unit_create_list.sort(key=len)
    # pprint(switch_create_list)
    # pprint(rnd_create_list)
    # pprint(unit_create_list)
    create_wwise_amb(rnd_create_list, switch_create_list)

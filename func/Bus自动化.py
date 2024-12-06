import os

from waapi import WaapiClient
import openpyxl
from pprint import pprint

import config
import module.cloudfeishu.cloudfeishu_h as cloudfeishu_h
import module.waapi.waapi_h as waapi_h
import module.oi.oi_h as oi_h
import module.excel.excel_h as excel_h

"""根目录获取"""
# 获取当前脚本的文件名及路径
script_name_with_extension = __file__.split('/')[-1]
# 去掉扩展名
script_name = script_name_with_extension.rsplit('.', 1)[0]
# 替换为files
root_path = script_name.replace("func", "files")
# print(f"当前脚本的名字是: {root_path}")

"""文件获取"""
# 规范检查表
excel_bus_path = os.path.join(root_path, config.excel_bus_path)
# print(wb)

"""在线表获取"""
# 规范检查表
check_name_token = config.bus_sheet_token
cloudfeishu_h.download_cloud_sheet(check_name_token, excel_bus_path)

"""功能数据初始化"""
# bus创建列表
bus_create_list = []
# bus unit创建列表
unit_create_list = []

# 获取xlsx的workbook
wb = openpyxl.load_workbook(excel_bus_path)
# 获取xlsx的所有sheet
sheet_names = wb.sheetnames

with WaapiClient() as client:
    """*****************Wwise功能区******************"""
    """设置对象属性"""


    def set_obj_property(obj_id, property, value):
        args = {
            "object": obj_id,
            "property": property,
            "value": value
        }
        client.call("ak.wwise.core.object.setProperty", args)
        # print(name_2 + ":" + str(trigger_rate))


    """obj创建"""


    def create_wwise_obj(type, parent, name, notes):
        unit_object = client.call("ak.wwise.core.object.create",
                                  waapi_h.args_object_create(parent,
                                                             type, name, notes))
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


    """Bus创建"""


    def create_wwise_bus(unit_list, bus_list):
        for bus_name in bus_list:
            # 需要创建unit的
            if bus_name in unit_list:
                # 查找已有unit列表
                unit_have_dict = get_wwise_type_list(config.wwise_bus_path, "WorkUnit")
                # pprint(unit_have_dict)
                # 查找已有bus列表
                bus_have_dict = get_wwise_type_list(config.wwise_bus_path, "Bus")
                # Wwise里没有unit的的
                if bus_name not in unit_have_dict:
                    # 查找该unit需要放的父级是谁
                    bus_name_parent = oi_h.find_longest_prefix_key(bus_name, unit_have_dict)
                    # 根目录创建
                    if not bus_name_parent:
                        oi_h.print_error(bus_name + "：未找到可以存放的父级Unit，请创建")
                    else:

                        # WorkUnit创建
                        unit_object = create_wwise_obj("WorkUnit",
                                                       bus_have_dict[bus_name_parent], bus_name,
                                                       "")
                        # Bus创建
                        bus_object = create_wwise_obj("Bus",
                                                      unit_object['id'], bus_name,
                                                      "")
            # 需要创建bus的
            else:
                # 查找已有bus列表
                bus_have_dict = get_wwise_type_list(config.wwise_bus_path, "Bus")
                if bus_name not in bus_have_dict:
                    # 查找该unit需要放的父级是谁
                    bus_name_parent = oi_h.find_longest_prefix_key(bus_name, bus_have_dict)
                    # Bus创建
                    bus_object = create_wwise_obj("Bus",
                                                  bus_have_dict[bus_name_parent], bus_name,
                                                  "")


    """获取需要创建的Bus列表"""


    def get_create_bus_name_list():
        # 遍历每一行
        for row in sheet.iter_rows():
            bus_name = None
            # 遍历每一列
            for cell in row:
                cell_value = excel_h.check_is_mergecell(cell, sheet)
                if cell_value:
                    is_unit = False
                    # print(cell_value)
                    if not bus_name:
                        bus_name = cell_value
                        is_unit = True
                    else:
                        bus_name += "_" + cell_value
                        fill = cell.fill
                        if fill.fill_type == 'solid':
                            if fill.start_color.index == "FFD9F3FD":
                                is_unit = True
                    if bus_name not in bus_create_list:
                        bus_create_list.append(bus_name)
                    if is_unit:
                        if bus_name not in unit_create_list:
                            unit_create_list.append(bus_name)

            # print()


    """*****************主程序处理******************"""

    # 查找已有bus unit列表

    # 加载所有工作表
    for sheet_name in sheet_names:
        sheet = wb[sheet_name]
        if sheet_name == "Bus创建":
            get_create_bus_name_list()
            # pprint(bus_create_list)
            # pprint(unit_create_list)
            # 列表排序
            bus_create_list.sort(key=len)
            unit_create_list.sort(key=len)
            create_wwise_bus(unit_create_list, bus_create_list)

            # 若自建Bus则置灰,颜色设为27
            # 查找已有bus列表
            bus_have_dict = get_wwise_type_list(config.wwise_bus_path, "Bus")
            for bus_have_name in bus_have_dict:
                if bus_have_name not in bus_create_list:
                    if (bus_have_name not in config.bus_no_color_list) and (
                    not oi_h.check_by_re(r'^Aux', bus_have_name)):
                        set_obj_property(bus_have_dict[bus_have_name], "OverrideColor", True)
                        set_obj_property(bus_have_dict[bus_have_name], "Color", 0)

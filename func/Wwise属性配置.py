import os

from waapi import WaapiClient
from pprint import pprint

import config
import module.cloudfeishu.cloudfeishu_h as cloudfeishu_h
import module.waapi.waapi_h as waapi_h
import module.oi.oi_h as oi_h

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
excel_check_path = os.path.join(root_path, config.excel_config_path)
# print(wb)

"""在线表获取"""
# 规范检查表
check_name_token = config.property_config_token
cloudfeishu_h.download_cloud_sheet(check_name_token, excel_check_path)

"""********************功能函数***********************"""

with WaapiClient() as client:
    """*****************Wwise功能区******************"""
    """修改Wwise内容的名字"""


    def change_name_by_wwise_content(obj_id, name, old_name, obj_type):
        args = {
            "objects": [
                {

                    "object": obj_id,
                    "name": name,
                }
            ]
        }
        client.call("ak.wwise.core.object.set", args)
        oi_h.print_warning(old_name + "(" + obj_type + ")改名为：" + name)


    """设置obj的notes"""


    def set_obj_notes(obj_id, notes_value):
        args = {
            'object': obj_id,
            'value': notes_value
        }
        client.call("ak.wwise.core.object.setNotes", args)
        # print_warning(obj_name + "描述更改为：" + notes_value)


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


    """通用内容创建"""


    def create_obj_content(state_group_parent_path, state_group_type,
                           state_group_name, state_group_desc):
        obj_info = {}
        flag = 0
        state_group_id = ""
        # 查找此state_group是否存在
        state_group_list = client.call("ak.wwise.core.object.get",
                                       waapi_h.waql_by_type(state_group_type, state_group_parent_path,
                                                            ),
                                       options=config.options)['return']
        # pprint(state_group_list)
        # state_group已存在
        for state_group_dict in state_group_list:
            if state_group_dict['name'] == state_group_name:
                flag = 1
                # 设置notes
                set_obj_notes(state_group_dict['id'], state_group_desc)
                state_group_id = state_group_dict['id']
                return state_group_id
        # state_group不存在，需要创建
        # 不存在则创建
        if flag == 0:
            obj_info = create_wwise_obj(state_group_type, state_group_parent_path, state_group_name, state_group_desc)
        return obj_info['id']


    """Bank内容放置"""


    def set_bank_content(bank_path, obj_path):
        return client.call("ak.wwise.core.soundbank.setInclusions", waapi_h.args_bank_create(bank_path, obj_path))


    """********************功能函数***********************"""
    """创建State Bank"""


    def create_state_bank(state_group_name, event_name, event_desc):

        """bank unit创建"""
        bank_parent_id = create_obj_content(config.wwise_state_bank_path, "WorkUnit", state_group_name,
                                            "")

        """bank创建"""
        # event命名
        bank_state_name = event_name.replace("AKE", "BNK")
        bank_state_id = create_obj_content(bank_parent_id, "SoundBank", bank_state_name,
                                           event_desc)
        return bank_state_id


    """State Bank自动划分"""


    def auto_divide_state_bank():
        # 获取有颜色标识的State Event
        state_bank_dict = {}
        wwise_event_list = client.call("ak.wwise.core.object.get",
                                       waapi_h.waql_by_type("State", config.wwise_state_path,
                                                            ),
                                       options=config.options)['return']
        for wwise_event_dict in wwise_event_list:
            # pprint(wwise_event_dict)
            if wwise_event_dict['Color'] == 3:
                # 查找引用State的对象
                refer_list = client.call("ak.wwise.core.object.get",
                                         waapi_h.waql_find_refer(wwise_event_dict['id'],
                                                                 ),
                                         options=config.options)['return']
                bank_id = None
                # 先创建bank
                for refer_dict in refer_list:
                    # 有引用的Event，Color需要设置为3
                    if refer_dict['type'] == 'Action':
                        set_obj_property(refer_dict['parent']['id'], "OverrideColor", True)
                        set_obj_property(refer_dict['parent']['id'], "Color", 3)
                        bank_id = create_state_bank(wwise_event_dict['parent']['name'], refer_dict['parent']['name'],
                                                    "")
                        set_bank_content(bank_id, refer_dict['parent']['id'])

                if bank_id:
                    # 查找需要分包的对象，即两个switch容器
                    for refer_dict in refer_list:
                        # 环境声switch
                        if refer_dict['type'] == 'SwitchContainer':
                            switch_children_list = client.call("ak.wwise.core.object.get",
                                                               waapi_h.waql_find_children(refer_dict['id'],
                                                                                          ),
                                                               options=config.options)['return']
                            for switch_children_dict in switch_children_list:
                                if wwise_event_dict['name'] in switch_children_dict['name']:
                                    # pprint(wwise_event_dict['name'])
                                    # pprint(switch_children_dict)
                                    # print()
                                    set_bank_content(bank_id, switch_children_dict['id'])
                                    break
                        # 音乐switch
                        if refer_dict['type'] == 'MusicSwitchContainer':
                            switch_children_list = client.call("ak.wwise.core.object.get",
                                                               waapi_h.waql_find_children(refer_dict['id'],
                                                                                          ),
                                                               options=config.options)['return']
                            for switch_children_dict in switch_children_list:
                                if wwise_event_dict['name'] in switch_children_dict['name']:
                                    # pprint(wwise_event_dict['name'])
                                    # pprint(switch_children_dict)
                                    # print()
                                    set_bank_content(bank_id, switch_children_dict['id'])
                                    break


    """*****************主程序处理******************"""
    # State Bank自动划分
    auto_divide_state_bank()

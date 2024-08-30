#!/usr/bin/env python3

from waapi import WaapiClient
from pprint import pprint
import os
import re

with WaapiClient() as client:
    """查找对象"""


    def find_obj(args):
        options = {
            'return': ['name', 'id', 'path']

        }
        obj_sub_list = client.call("ak.wwise.core.object.get", args, options=options)['return']
        if not obj_sub_list:
            obj_sub_id = ""
            obj_sub_path = ""
        else:
            obj_sub_id = obj_sub_list[0]['id']
            obj_sub_path = obj_sub_list[0]['path']
        return obj_sub_list, obj_sub_id, obj_sub_path


    selected_list = client.call("ak.wwise.ui.getSelectedObjects")['objects']
    # pprint(selected)
    # [{'id': '{24279D4F-74C7-45D4-B286-D6FF49DE1605}',
    #   'name': 'Char_Skill_C02_Atk5'},
    #  {'id': '{B56B4DDE-5395-482B-9632-80EB6966D8BF}',
    #   'name': 'Char_Skill_C02_Atk5_End'},
    #  {'id': '{18E474C3-4F09-4A43-ACD0-4FC166721076}',
    #   'name': 'Char_Skill_C02_Battle_Exit'}]

    if selected_list:
        # 在子对象的父级下创建父级对象
        # 1.查找Actor-Mixer
        _, parent_id, _ = find_obj(
            {'waql': ' "%s"  select parent ' % selected_list[0]['id']})
        if parent_id:
            # 定义一个lambda函数，用于获取字典中某个键的值
            get_selected_one_name = lambda dict_selected: dict_selected['name']
            # 针对一个全是字典元素的列表，对其取某个键的值并组成一个新列表
            selected_name_list = list(map(get_selected_one_name, selected_list))
            # pprint(selected_name_list)
            # ['Char_Skill_C02_Atk5',
            #  'Char_Skill_C02_Atk5_End',
            #  'Char_Skill_C02_Atk_Rush',
            #  'Char_Skill_C02_Battle_Enter1',
            #  'Char_Skill_C02_Battle_Exit']

            # 取出所选元素名称的共同前缀,作为父级switch容器的名称
            switch_name = os.path.commonprefix(selected_name_list)
            # 如果最后为_，则去掉
            switch_name = re.sub(r"(_)$", "", switch_name)

            if switch_name:
                # print(switch_name)
                # 2.创建Switch父级容器
                args = {
                    "objects": [
                        {
                            "onNameConflict": "rename",
                            "object": parent_id,
                            "children": [
                                {
                                    "type": "SwitchContainer",
                                    "name": switch_name
                                }
                            ]

                        }
                    ]
                }
                result = client.call("ak.wwise.core.object.set", args)
                # print(result)
                # {'objects': [{'children': [{'id': '{3BC5DF47-A8EF-4B64-A9BC-029E852E9A0C}',
                #                             'name': 'Char_Skill_C02_Atk'}],
                #               'id': '{9E148809-7E30-4D78-B7EF-BE2783053FCF}',
                #               'name': 'Char_Skill_C02'}]}
                switch_id = result['objects'][0]['children'][0]['id']
                if switch_id:
                    # print(switch_id)
                    # 移动子对象到父级对象下
                    get_selected_one_id = lambda dict_selected: dict_selected['id']
                    selected_id_list = list(map(get_selected_one_id, selected_list))
                    for selected_id in selected_id_list:
                        args = {
                            "object": selected_id,
                            "parent": switch_id,
                            "onNameConflict": "fail"
                        }

                        client.call("ak.wwise.core.object.move", args)

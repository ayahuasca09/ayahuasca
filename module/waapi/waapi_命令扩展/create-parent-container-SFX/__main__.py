#!/usr/bin/env python3

from waapi import WaapiClient
import os
import re

if __name__ == '__main__':

    # 容器类型选择
    container_type = ""


    def print_log(warn_info):
        print("[log]" + warn_info)


    with WaapiClient() as client:

        # 撤销开始
        client.call("ak.wwise.core.undo.beginGroup")
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


        """获取列表中某种元素的值，并组成一个新列表"""


        # dict_type：要获取的字典值
        # refer_list：需要传入的原列表
        def create_one_element_list(dict_type, refer_list):
            # 定义一个lambda函数，用于获取字典中某个键的值
            get_refer_action_id = lambda dict_refer: dict_refer[dict_type]
            # 针对一个全是字典元素的列表，对其取某个键的值并组成一个新列表
            refer_id_list = list(map(get_refer_action_id, refer_list))
            return refer_id_list


        """******************主程序******************"""

        # 获取选中的对象
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
            _, mixer_id, _ = find_obj(
                {'waql': ' "%s"  select parent ' % selected_list[0]['id']})
            if mixer_id:
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
                parent_name = os.path.commonprefix(selected_name_list)
                # 如果最后为_，则去掉
                parent_name = re.sub(r"(_)$", "", parent_name)

                if parent_name:
                    # print(parent_name)
                    # 2.创建父级容器
                    # 选择要转换的容器类型
                    LUT = ['SwitchContainer',
                           ['RandomSequenceContainer', 0, '@SequenceContainer'],
                           'BlendContainer']
                    print("--------------------------")
                    print("输入数字对应容器类型如下：")
                    for i in range(len(LUT)):
                        print("{}、".format(i + 1) + str(LUT[i]))
                    print("--------------------------")

                    flag = 1
                    while flag:
                        temp = input("请输入要转换为的容器类型:")
                        if temp.isdigit() != True:
                            print("请输入1~6之间的整数")
                            continue
                        temp = int(temp) - 1
                        if temp < 6:
                            flag = 0
                        else:
                            print("请输入1~6之间的整数")

                    # container_type是LUT中元素的索引
                    container_type = temp

                    print("")

                    # 输入5得到的container_type值：4

                    # 判断LUT中的元素是否为str，不是则走另一种生成方式
                    print_type = ""
                    if isinstance(LUT[container_type], str):
                        print_type = LUT[container_type]
                        args = {
                            "objects": [
                                {
                                    "onNameConflict": "rename",
                                    "object": mixer_id,
                                    "children": [
                                        {
                                            "type": LUT[container_type],
                                            "name": parent_name
                                        }
                                    ]

                                }
                            ]
                        }
                    else:
                        print_type = LUT[container_type][2]
                        args = {
                            "objects": [
                                {
                                    "onNameConflict": "rename",
                                    "object": mixer_id,
                                    "children": [
                                        {
                                            "type": LUT[container_type][0],
                                            "name": parent_name,
                                            "@RandomOrSequence": LUT[container_type][1]
                                        }
                                    ]

                                }
                            ]
                        }

                    if args:
                        result = client.call("ak.wwise.core.object.set", args)
                        # print(result)
                        # {'objects': [{'children': [{'id': '{3BC5DF47-A8EF-4B64-A9BC-029E852E9A0C}',
                        #                             'name': 'Char_Skill_C02_Atk'}],
                        #               'id': '{9E148809-7E30-4D78-B7EF-BE2783053FCF}',
                        #               'name': 'Char_Skill_C02'}]}
                        parent_id = result['objects'][0]['children'][0]['id']
                        if parent_id:
                            print_log(print_type + "容器创建：" + parent_name)
                            # print(parent_id)
                            # 移动子对象到父级对象下
                            for selected_dict in selected_list:
                                args = {
                                    "object": selected_dict['id'],
                                    "parent": parent_id,
                                    "onNameConflict": "fail"
                                }

                                client.call("ak.wwise.core.object.move", args)

                                """删除选中容器引用的事件"""
                                # 查找选中的容器引用的Event Action
                                args = {
                                    'waql': '"%s" select referencesTo' % selected_dict['id']
                                }
                                refer_list, refer_id, _ = find_obj(args)
                                # print(refer_id)
                                # 定义一个lambda函数，用于获取字典中某个键的值
                                get_refer_action_id = lambda dict_refer: dict_refer['id']
                                # 针对一个全是字典元素的列表，对其取某个键的值并组成一个新列表
                                refer_id_list = list(map(get_refer_action_id, refer_list))
                                # print(refer_id_list)

                                # 查找Event Action的Event
                                for refer_id in refer_id_list:
                                    args = {
                                        'waql': '"%s" select parent' % refer_id
                                    }
                                    event_list, _, _ = find_obj(args)

                                    # 删除主要以此选中对象创建的事件
                                    # 此操作排除了此对象夹在其他事件中，但不以此对象相关命名的事件
                                    for event_dict in event_list:
                                        if selected_dict['name'] in event_dict['name']:
                                            print_log("Event删除：" + event_dict['name'])
                                            args = {
                                                "object": "%s" % event_dict['id']
                                            }
                                            client.call("ak.wwise.core.object.delete", args)

        client.call("ak.wwise.core.undo.endGroup", displayName="parent创建撤销")

        print("")
        input("按任意字符结束")

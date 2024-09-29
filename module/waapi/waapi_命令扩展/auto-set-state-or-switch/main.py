from waapi import WaapiClient
from pprint import pprint

with WaapiClient() as client:
    # 撤销开始
    client.call("ak.wwise.core.undo.beginGroup")

    """报错捕获"""


    def print_error(error_info):
        global is_pass
        is_pass = False
        print("[error]" + error_info)


    def print_log(warn_info):
        print("[log]" + warn_info)


    """查找对象"""


    def find_obj(args):
        options = {
            'return': ['name', 'id', 'path', 'type', 'SwitchGroupOrStateGroup']

        }
        obj_sub_list = client.call("ak.wwise.core.object.get", args, options=options)['return']
        if not obj_sub_list:
            obj_sub_id = ""
            obj_sub_path = ""
        else:
            obj_sub_id = obj_sub_list[0]['id']
            obj_sub_path = obj_sub_list[0]['path']
        return obj_sub_list, obj_sub_id, obj_sub_path


    # 设置子集switch
    def set_switch_child(obj_child, switch_child):
        args = {
            'child': obj_child,
            'stateOrSwitch': switch_child
        }
        client.call("ak.wwise.core.switchContainer.addAssignment", args)


    # 获取选定对象
    objects_list = client.call("ak.wwise.ui.getSelectedObjects")['objects']
    # [{'id': '{CD4394F8-0E13-4556-9EBA-44B9194EDDA2}', 'name': 'Char_Skill_C03_Ult'}]

    for object_dict in objects_list:
        object_name = object_dict['name']
        object_id = object_dict['id']
        args = {
            'waql': '"%s" ' % object_id
        }
        object_info, _, _ = find_obj(args)
        # 查看是否为switch container
        if 'SwitchGroupOrStateGroup' in object_info[0]:
            # 是否指派group，若未指派则报错
            if object_info[0]['SwitchGroupOrStateGroup']['id'] == '{00000000-0000-0000-0000-000000000000}':
                print_error(object_name + ":请先指派Switch Group及默认值")
            else:
                # pprint(object_info)
                # 获取switch container的子容器
                args = {
                    'waql': '"%s" select children' % object_id
                }
                object_child_list, _, _ = find_obj(args)

                # 获取SwitchGroupOrStateGroup的子级
                args = {
                    'waql': '"%s" select children' % object_info[0]['SwitchGroupOrStateGroup']['id']
                }
                state_child_list, _, _ = find_obj(args)

                # pprint(object_child_list)
                # 查找state是否与子级名称匹配
                for object_child_dict in object_child_list:
                    flag = 0
                    for state_child_dict in state_child_list:
                        if state_child_dict['name'] in object_child_dict['name']:
                            # print(object_child_dict['name'])
                            set_switch_child(object_child_dict['id'], state_child_dict['id'])
                            print_log(object_child_dict['name'] + "映射State/Switch值为：" + state_child_dict['name'])
                            flag = 1
                            break
                    if flag == 0:
                        print_log(object_child_dict['name'] + "未找到相应的State/Switch")

    client.call("ak.wwise.core.undo.endGroup", displayName="State/Switch指派撤销")

    print("")
    input("按任意字符结束")

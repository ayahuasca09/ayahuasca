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


    """查找对象"""


    def find_obj(args):
        options = {
            'return': ['name', 'id', 'path', 'notes', 'originalWavFilePath']

        }
        obj_sub_list = client.call("ak.wwise.core.object.get", args, options=options)['return']
        if not obj_sub_list:
            obj_sub_id = ""
            obj_sub_path = ""
        else:
            obj_sub_id = obj_sub_list[0]['id']
            obj_sub_path = obj_sub_list[0]['path']
        return obj_sub_list, obj_sub_id, obj_sub_path


    """创建事件的参数"""


    def get_event_args(event_parent_path, event_name, event_action, event_target, notes):
        args_new_event = {
            # 上半部分属性中分别为 Event 创建后存放的路径、类型、名称、遇到名字冲突时的处理方法
            "parent": event_parent_path,
            "type": "Event",
            "name": event_name,
            "onNameConflict": "rename",
            "children": [
                {
                    # 为其创建播放行为，名字留空，使用 @ActionType 定义其播放行为为 Play，@Target 为被播放的声音对象
                    "name": "",
                    "type": "Action",
                    "@ActionType": event_action,
                    "@Target": event_target
                }
            ],
            "notes": notes
        }
        return args_new_event


    """获取wwise对象的最长字符串父级"""


    def find_obj_parent(obj_name, waql):
        # 所有Unit获取
        parent_list, parent_id, _ = find_obj(waql)

        # 存储符合的对象的父级，最后找出最长最符合的
        parent_len = 0
        obj_parent_id = {}
        if parent_list != None:
            for parent_dict in parent_list:
                if parent_dict['name'] in obj_name:
                    if len(parent_dict['name']) > parent_len:
                        parent_len = len(parent_dict['name'])
                        obj_parent_id = parent_dict['id']

        return obj_parent_id


    """********************主程序********************"""
    # 获取选中对象
    selected_list = client.call("ak.wwise.ui.getSelectedObjects")['objects']
    # pprint(selected)
    # [{'id': '{24279D4F-74C7-45D4-B286-D6FF49DE1605}',
    #   'name': 'Char_Skill_C02_Atk5'}]
    selected_name = selected_list[0]['name']
    selected_id = selected_list[0]['id']
    media_list, media_id, _ = find_obj(
        {'waql': ' "%s"  ' % selected_id})
    selected_notes = media_list[0]['notes']
    # print(selected_notes)

    if selected_list:
        if len(selected_list) > 1:
            print_error("只能为一个对象创建事件")
        else:
            # 父级路径查找
            parent_id = find_obj_parent(selected_name, {
                'waql': ' "\\Events\\v1" select descendants where type = "WorkUnit" '})
            event_name = "AKE_" + "Play_" + selected_name
            event_args = get_event_args(parent_id, event_name, 1, selected_id, selected_notes)
            new_event = client.call("ak.wwise.core.object.create", event_args)
            new_event_name = new_event['name']
            # print(new_event_name)
            print("Play Event创建：" + new_event_name)
            if "_LP" in selected_name:
                event_name = "AKE_" + "Stop_" + selected_name
                event_args = get_event_args(parent_id, event_name, 2, selected_id, selected_notes + "停止")
                new_event = client.call("ak.wwise.core.object.create", event_args)
                new_event_name = new_event['name']
                print("Stop Event创建：" + new_event_name)
            print("")
            input("按任意字符结束")

    client.call("ak.wwise.core.undo.endGroup", displayName="Event创建撤销")

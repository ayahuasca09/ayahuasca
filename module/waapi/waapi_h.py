"""获取对象信息"""


def find_obj(obj_list):
    if not obj_list:
        obj_sub_id = ""
        obj_sub_path = ""
    else:
        obj_sub_id = obj_list[0]['id']
        obj_sub_path = obj_list[0]['path']
    return obj_list, obj_sub_id, obj_sub_path


"""WAQL查询"""


# 通过type查询
def waql_by_type(waql_type, obj_id_or_path):
    return {'waql': ' "%s" select descendants where type = "%s" ' % (obj_id_or_path, waql_type)}


# 通过类型查询所有
def waql_from_type(waql_type):
    return {'waql': f'from type  {waql_type} '}


# 通过ID查询对象
def waql_from_id(waql_id):
    return {'waql': 'from object  "%s" ' % waql_id}


"""通用内容创建"""


def create_obj_content(state_group_parent_path, state_group_type,
                       state_group_name, state_group_desc, state_id):
    flag = 0
    state_group_id = ""
    state_group_path = ""
    # 查找此state_group是否存在
    state_group_list, _, _ = find_obj(
        {'waql': ' "%s" select descendants where type = "%s"' % (
            state_group_parent_path, state_group_type)})
    # pprint(state_group_list)
    # state_group已存在
    for state_group_dict in state_group_list:
        if state_group_dict['name'] == state_group_name:
            flag = 1
            # 设置notes
            set_obj_notes(state_group_dict['id'], state_group_desc)
            state_group_id = state_group_dict['id']
            state_group_path = state_group_dict['path']
            # 设置bpm
            if (state_group_dict['type'] == "MusicSegment") and bpm_value and bpm_value != 0:
                set_obj_property(state_group_dict['id'], "OverrideClockSettings", True)
                set_obj_property(state_group_dict['id'], "Tempo", bpm_value)
            break
        # 改名了但描述一致
        elif state_group_dict['notes'] == state_group_desc:
            flag = 2
            # 改名
            change_name_by_wwise_content(state_group_dict['id'], state_group_name, state_group_dict['name'],
                                         state_group_type)
            state_group_id = state_group_dict['id']
            state_group_path = state_group_dict['path']
            break
    # state_group不存在，需要创建
    # print(flag)
    # 不存在则创建
    if flag == 0:
        if state_group_type == "Event":
            # 专属trig用
            args = {
                # 选择父级
                "parent": state_group_parent_path,
                # 创建类型名称
                "type": state_group_type,
                "name": state_group_name,
                "notes": state_group_desc,
                "children": [
                    {
                        "name": "",
                        "type": "Action",
                        "@ActionType": 35,
                        "@Target": state_id
                    }]
            }
        else:
            # 创建的rnd属性
            args = {
                # 选择父级
                "parent": state_group_parent_path,
                # 创建类型名称
                "type": state_group_type,
                "name": state_group_name,
                "notes": state_group_desc,
            }
        state_group_object = client.call("ak.wwise.core.object.create", args)
        # print(state_group_object)
        state_group_id = state_group_object['id']
        _, _, state_group_path = find_obj(
            {'waql': ' "%s" ' % state_group_id})
        print_warning(state_group_object['name'] + ":" + state_group_type + "已创建")

    return state_group_id, state_group_path

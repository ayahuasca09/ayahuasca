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


# 查找引用
def waql_find_refer(waql_id):
    return {'waql': '"%s" select referencesTo' % waql_id}


# 查找子级
def waql_find_children(waql_id):
    return {'waql': '"%s" select children' % waql_id}


"""args模板"""


# 创建obj的模板
def args_object_create(parent, type, name, notes):
    args = {
        # 选择父级
        "parent": parent,
        # 创建类型名称
        "type": type,
        "name": name,
        "notes": notes,
        "onNameConflict": "replace"
    }
    return args


# 创建rnd的模板
def args_rnd_create(parent, type, name, notes):
    args = {
        # 选择父级
        "parent": parent,
        # 创建类型名称
        "type": type,
        "name": name,
        "notes": notes,
        "onNameConflict": "replace",
        "@RandomOrSequence": 1,
        "@NormalOrShuffle": 1,
        "@RandomAvoidRepeatingCount": 1
    }
    return args


# unit_object = client.call("ak.wwise.core.object.create", args)
#             # {'id': '{70821958-7177-49F5-A442-D0026B5C47C4}', 'name': 'Char_Mov_Gen_SHeel'}

# 创建bank的模板
def args_bank_create(bank_path, obj_path):
    args = {
        "soundbank": bank_path,
        "operation": "add",
        "inclusions": [
            {
                # 此处为 Just_a_Bank 内要放入的内容，filter 是要把哪些数据放进 SoundBank
                "object": obj_path,
                "filter": [
                    "events",
                    "structures",
                    "media"
                ]
            }
        ]
    }
    return args

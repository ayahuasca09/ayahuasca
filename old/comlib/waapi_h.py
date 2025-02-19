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


# 查找某个对象的RTPC
def waql_find_RTPC(obj_id_or_path):
    return {'waql': ' "%s" select RTPC  ' % obj_id_or_path}


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


# 设置对象属性的模板
def args_set_obj_property(obj_id, property, value):
    args = {
        "object": obj_id,
        "property": property,
        "value": value
    }
    # client.call("ak.wwise.core.object.setProperty", args)
    # print(name_2 + ":" + str(trigger_rate))
    return args


# args=waapi_h.args_set_obj_property(obj_id, "OverrideColor", False)
# client.call("ak.wwise.core.object.setProperty", args)


# 设置obj的notes
def args_set_obj_notes(obj_id, notes_value):
    args = {
        'object': obj_id,
        'value': notes_value
    }
    # client.call("ak.wwise.core.object.setNotes", args)
    # print_warning(obj_name + "描述更改为：" + notes_value)
    return args


# args=waapi_h.args_set_obj_notes(obj_dict['id'], obj_desc)
# client.call("ak.wwise.core.object.setNotes", args)


# 修改Wwise内容的名字
def args_change_name_by_wwise_content(obj_id, name, old_name, obj_type):
    args = {
        "objects": [
            {

                "object": obj_id,
                "name": name,
            }
        ]
    }
    # client.call("ak.wwise.core.object.set", args)
    # oi_h.print_warning(old_name + "(" + obj_type + ")改名为：" + name)
    return args


# args = waapi_h.args_change_name_by_wwise_content(obj_dict['id'], obj_name, obj_dict['name'],
#                                                                  obj_type)
# client.call("ak.wwise.core.object.set", args)


# ducking rtpc模板
def args_rtpc_ducking_create(obj_id, rtpc_id):
    args = {
        "objects": [
            {
                "object": obj_id,
                "@RTPC": [
                    {
                        "type": "RTPC",
                        "name": "",
                        "@Curve": {
                            "type": "Curve",
                            "points": [
                                {
                                    "x": -48,
                                    "y": 0,
                                    "shape": "Linear"
                                },
                                {
                                    "x": 0,
                                    "y": -6,
                                    "shape": "Linear"
                                }
                            ]
                        },
                        # "notes": "a new rtpc",
                        "@PropertyName": "OutputBusVolume",
                        "@ControlInput": rtpc_id
                        # ControlInput填的是rtpc的id
                    }
                ]
            }

        ]
    }
    return args


# wwise_info_list = client.call("ak.wwise.core.object.set", args)
# pprint(wwise_info_list)
# RTPC的类型就是RTPC，id是@RTPC的id
# {'objects': [{'@RTPC': [{'id': '{368C32DE-C9E1-4950-ABA3-78E33FF85CAF}',
#                          'name': ''}],
#               'id': '{02AA9D36-C4CB-45D4-B0EC-9B17749258E3}',
#               'name': 'Mon'}]}

# 设置effect的模板
def args_effect_set(obj_id, effect_id):
    args = {
        "objects": [
            {
                "object": obj_id,
                "@Effect0": effect_id
            }
        ]
    }
    return args


# {
#     "objects": [
#         {
#             "object": "\\Actor-Mixer Hierarchy\\Default Work Unit\\MySound",
#             "@Effects": [
#                 {
#                     "type": "EffectSlot",
#                     "name": "",
#                     "@Effect": {
#                         "type": "Effect",
#                         "name": "myCustomEffect",
#                         "classId": 7733251,
#                         "@PreDelay": 24,
#                         "@RoomShape": 99
#                     }
#                 }
#             ]
#         }
#     ]
# }


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


# 创建effects的模板
def args_effect_create(parent, type, name):
    args = {
        "objects": [
            {
                "object": parent,
                "children": [
                    {
                        "type": "Effect",
                        "name": name,
                        "classId": type,
                    }
                ]
            }
        ],
        "onNameConflict": "merge"

    }
    return args


# 参考代码
# #!/usr/bin/env python3
# from waapi import WaapiClient, CannotConnectToWaapiException
# from pprint import pprint
#
# try:
#     # Connecting to Waapi using default URL
#     with WaapiClient() as client:
#         args = {
#             "objects": [
#                 {
#                     "object": "\\Effects\\Amb",
#                     "children": [
#                         {
#                             "type": "Effect",
#                             "name": "MyTestEffect",
#                             "classId": 8454147, # Meter
#                             "@AttackTime": 10,
#                         }
#                     ]
#                 }
#             ],
#             "onNameConflict": "merge"
#         }
#
#
#         result = client.call("ak.wwise.core.object.set", args)
#         pprint(result)
#
# except CannotConnectToWaapiException:
#     print("Could not connect to Waapi: Is Wwise running and Wwise Authoring API enabled?")

# 设置stategroup的模板
def args_stategroup_set(obj_id, stategroup_id):
    args = {
        "object": obj_id,
        "stateGroups": [
            stategroup_id
        ]
    }
    return args


# unit_object = client.call("ak.wwise.core.object.setStateGroups", args)


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


# 音乐媒体资源导入的模板
def args_mus_create(source_path, obj_id, sub_path):
    args = {
        # createNew
        # useExisting：会增加一个新媒体文件但旧的不会删除
        # replaceExisting:会销毁Sound，上面做的设置都无了
        "importOperation": "replaceExisting",
        "imports": [
            {
                "audioFile": source_path,
                "objectPath": obj_id,
                "originalsSubFolder": sub_path
            }
        ]
    }
    return args


# client.call("ak.wwise.core.audio.import", args_import)

# sfx媒体资源导入的模板
def args_sfx_create(source_path, rnd_path, media_name, sub_path):
    args = {
        "importOperation": "replaceExisting",
        "default": {
            "importLanguage": "SFX"
        },
        "imports": [
            {
                "audioFile": source_path,
                "objectPath": rnd_path + '\\<Sound SFX>' + media_name,
                "originalsSubFolder": sub_path
                #                                                         名为Test 0的顺序容器            名为My SFX 0 的音效
                # "objectPath": "\\Actor-Mixer Hierarchy\\Default Work Unit\\<Sequence Container>Test 0\\<Sound SFX>My SFX 0"
            }
        ]
    }
    return args

# client.call("ak.wwise.core.audio.import", args_import)

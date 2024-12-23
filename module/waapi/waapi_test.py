from waapi import WaapiClient
from pprint import pprint

"""跨脚本调用测试"""
# args = {
#     'waql': '$ "\\Events" select descendants where type = "Event" '
# }
# obj_sub_list, obj_sub_id, obj_sub_path = waapi_h.find_obj(args)
# print(obj_sub_list)


with WaapiClient() as client:
    def find_obj(args):
        options = {
            'return': ['name', 'id', 'notes', 'originalWavFilePath', 'isIncluded', 'IsLoopingEnabled',
                       'musicPlaylistRoot', 'LoopCount', 'PlaylistItemType', 'owner', 'parent', 'type', 'Sequences']

        }
        obj_sub_list = client.call("ak.wwise.core.object.get", args, options=options)['return']
        if not obj_sub_list:
            obj_sub_id = ""
        else:
            obj_sub_id = obj_sub_list[0]['id']
        return obj_sub_list, obj_sub_id


    # obj_sub_list = client.call("ak.wwise.core.object.get", args, options=config.options)['return']

    """soundbank生成测试"""
    # args = {
    #     "soundbanks": [
    #         {"name": "ababa"}
    #     ],
    #     "writeToDisk": True
    # }
    #
    # client.call("ak.wwise.core.soundbank.generate", args)

    """基本功能测试"""
    # args = {
    #     'waql': '$ "\\Events" select descendants where type = "Event" '
    # }
    # options = {
    #     'return': ["parent", "name"]
    # }
    # # 存储了所有Event的单元结构
    # wwise_info_list = client.call("ak.wwise.core.object.get", args, options = options)['return']
    # pprint(wwise_info_list)

    """获取选中对象的打包流媒体的路径"""
    # objects_list = client.call("ak.wwise.ui.getSelectedObjects")['objects']
    # obj_id = objects_list[0]['id']
    # print(obj_id)
    # args = {
    #     'waql': '"%s"' % obj_id
    # }
    # options = {
    #     'return': ['name','convertedWemFilePath']
    # }
    # # 存储了所有Event的单元结构
    # wwise_info_list = client.call("ak.wwise.core.object.get", args, options=options)['return']
    # pprint(wwise_info_list)

    """文件导入测试"""
    # args_import = {
    #     # createNew
    #     # useExisting：会增加一个新媒体文件但旧的不会删除
    #     # replaceExisting:会销毁Sound，上面做的设置都无了
    #     "importOperation": "replaceExisting",
    #     "default": {
    #         "importLanguage": "SFX"
    #     },
    #     "imports": [
    #         {
    #             "audioFile": r"F:\pppppy\SP\module\waapi\waapi_音频资源导入自动化\Media_Temp.wav",
    #             "objectPath": '\\Actor-Mixer Hierarchy\\v1\\Char\\Char\\Char_Skill\\Char_Skill\\Char_Skill_C01\\Char_Skill_C01\\Char_Skill_C01_Strafe_Enter\\<Sound SFX>My SFX 0',
    #             "originalsSubFolder": "Char"
    #             #                                                         名为Test 0的顺序容器            名为My SFX 0 的音效
    #             # "objectPath": "\\Actor-Mixer Hierarchy\\Default Work Unit\\<Sequence Container>Test 0\\<Sound SFX>My SFX 0"
    #         }
    #     ]
    # }
    # # 定义返回结果参数，让其只返回 Windows 平台下的信息，信息中包含 GUID 和新创建的对象名
    # opts = {
    #     "platform": "Windows",
    #     "return": [
    #         "id", "name"
    #     ]
    # }
    # client.call("ak.wwise.core.audio.import", args_import, options=opts)

    """对象删除测试"""
    # objects_list = client.call("ak.wwise.ui.getSelectedObjects")['objects']
    # for object_dict in objects_list:
    #     object_id = object_dict['id']
    #     args = {
    #         "object": "%s" % object_id
    #     }
    #     client.call("ak.wwise.core.object.delete", args)

    """查找对象引用测试"""
    #
    # def find_obj(args):
    #     options = {
    #         'return': ['name', 'id', 'notes']
    #
    #     }
    #     obj_sub_list = client.call("ak.wwise.core.object.get", args, options=options)['return']
    #     if not obj_sub_list:
    #         obj_sub_id = ""
    #     else:
    #         obj_sub_id = obj_sub_list[0]['id']
    #     return obj_sub_list, obj_sub_id
    #
    #
    # args = {
    #     'waql': '"%s" select referencesTo' % (
    #         '{2F8B7868-B380-4767-9681-1665DACC4087}')
    # }
    # refer_list, refer_id = find_obj(args)
    # print(refer_list)

    """设置对象notes"""
    # args = {
    #     'object': '{15E5DC76-6CD3-4051-86C2-4DA23B16EB13}',
    #     'value': "aaaaa"
    # }
    # client.call("ak.wwise.core.object.setNotes", args)

    # """设置对象属性/媒体资源路径修改测试"""
    # args = {
    #     "objects": [
    #         {
    #             # ______使用媒体对象SFX作为对象______
    #             # 修改媒体对象的命名,音量
    #             "object": "{C1466E57-C6EA-4798-B41A-EBE116DB3EFF}",
    #             # 还可以使用GUID作为路径
    #             # "object":"{B6D45483-A1BF-4EEC-87CF-2E0CD493B2CB}",
    #             "name": "a_02",
    #             "@Volume": -18,
    #             "@Pitch": 200,
    #             "notes": "asda"
    #
    #         }
    #     ]
    # }
    # options = {
    #     'return': ['name']
    # }
    # result = client.call("ak.wwise.core.object.set", args, options=options)
    # print(result)

    """external source自动转码测试"""
    # args = {
    #     "sources": [
    #         {
    #             "input": "F:\\pppppy\\SP\\module\\waapi\\waapi_Auto_Import_ExternalSource\\ExternalSource.wsources",
    #             "platform": "Windows",
    #             "output": 'S:\\Ver_1.0.0\\Project\\Content\\Audio\\GeneratedExternalSources\\Windows'
    #         },
    #         {
    #             "input": "F:\\pppppy\\SP\\module\\waapi\\waapi_Auto_Import_ExternalSource\\ExternalSource.wsources",
    #             "platform": "Android",
    #             "output": 'S:\\Ver_1.0.0\\Project\\Content\\Audio\\GeneratedExternalSources\\Android'
    #         },
    #         {
    #             "input": "F:\\pppppy\\SP\\module\\waapi\\waapi_Auto_Import_ExternalSource\\ExternalSource.wsources",
    #             "platform": "iOS",
    #             "output": 'S:\\Ver_1.0.0\\Project\\Content\\Audio\\GeneratedExternalSources\\iOS'
    #         }
    #     ]
    # }
    #
    # gen_log = client.call("ak.wwise.core.soundbank.convertExternalSources", args)

    """语音导入测试"""
    # args = {
    #     "importOperation": "createNew",
    #     "imports": [
    #         {
    #             "audioFile": "C:\\Users\\happyelements\\Desktop\\Rosy\\WILD\\VO_C10_01_Menu.wav",
    #             "objectPath": "\\Actor-Mixer Hierarchy\\v1\\VO\\<Sound Voice>hello\\<AudioFileSource>hello_cn",
    #             "importLanguage": "Chinese"
    #         }
    #     ]
    # }
    # client.call("ak.wwise.core.audio.import", args)

    """获取语音属性"""

    #
    # def find_obj(args):
    #     options = {
    #         'return': ['name', 'id', 'notes', 'originalWavFilePath', 'isIncluded']
    #
    #     }
    #     obj_sub_list = client.call("ak.wwise.core.object.get", args, options=options)['return']
    #     if not obj_sub_list:
    #         obj_sub_id = ""
    #     else:
    #         obj_sub_id = obj_sub_list[0]['id']
    #     return obj_sub_list, obj_sub_id
    #
    #
    # args = {
    #     'waql': '"%s" select children ' % (
    #         '{D2EDE786-525F-42F3-B31F-C5531798CFCA}')
    # }
    # refer_list, refer_id = find_obj(args)
    # pprint(refer_list)

    """获取事件引用的媒体是否为循环"""

    # args = {
    #     'waql': '"%s" select descendants where type= "Action" ' % (
    #         '{05244502-61D1-4B75-B7E1-5DD4852672D1}')
    # }
    #
    # options = {
    #     'return': ['Target', 'ActionType']
    #
    # }
    # obj_sub_list = client.call("ak.wwise.core.object.get", args, options=options)['return']
    # # pprint(obj_sub_list)
    #
    # # [{'ActionType': 1,
    # #   'Target': {'id': '{913295DF-370C-47EB-98B6-11E12815D7C7}',
    # #              'name': 'VO_Game_Battle_B01_09'}}]
    #
    # # 需要type=1或type=23或type=22才继续往下判断是否为lp
    # # 22	/States/Set State
    # # 23	/Set Switch
    # # 1	/Play
    # if obj_sub_list:
    #     for obj_dict in obj_sub_list:
    #         if obj_dict['ActionType'] == 1:
    #             # print("nice")
    #             args = {
    #                 'waql': '"%s" select descendants where type= "Sound" ' % (
    #                     obj_dict['Target']['id'])
    #             }
    #             sound_list, sound_id = find_obj(args)
    #             if sound_list:
    #                 # pprint(sound_list)
    #                 for sound_dict in sound_list:
    #                     if sound_dict['IsLoopingEnabled']==True:
    #                         # return True
    #                         break

    """筛选出所有有多个随机资源的音效"""
    # args = {
    #     'waql': '"%s" select descendants where type= "Sound" and name=/R\d+$/ and name=/^(?!VO).*/' % (
    #         '{5CB72D1B-8D4B-484D-8B9D-FC52BD496843}')
    # }
    # refer_list, refer_id = find_obj(args)
    # # pprint(refer_list)
    # for refer_dict in refer_list:
    #     print(refer_dict['name'])

    """查找音乐是否循环"""
    # # 先查找MusicPlaylistContainer容器
    # args = {
    #     'waql': '"%s" select descendants where type= "MusicPlaylistContainer" ' % (
    #         '{E8A5F6A0-AB3E-4166-BB68-C7D7ECC92B29}')
    # }
    # refer_list, refer_id = find_obj(args)
    # # pprint(refer_list)
    # for refer_dict in refer_list:
    #     # print(refer_dict['id'])
    #     # Mus_Login
    #     # Mus_Loading
    #     # Mus_Map_A02_Combat_Boss_Start
    #     # Mus_Map_A02_Combat_Boss_Stage1
    #     pprint(refer_dict['name'])
    #     args = {
    #         'waql': '"%s" ' % (
    #             refer_dict['musicPlaylistRoot']['id'])
    #     }
    #     options = {
    #         'return': ['name', 'LoopCount']
    #
    #     }
    #     obj_sub_list = client.call("ak.wwise.core.object.get", args, options=options)['return']
    #     pprint(obj_sub_list)

    """查找音乐列表下的item是否循环"""
    # PS：MusicPlaylistItem的根节点只有owner，
    # MusicPlaylistItem的其他子节点只有parent，parent为根节点
    # args = {
    #     'waql': 'from type MusicPlaylistItem '
    # }
    # musicplaylistitem_list, _ = find_obj(args)
    # # pprint(refer_list)
    #
    # # 使用一个列表保存为loop的musicplaylist container id
    # loop_musicplaylist_container_id_list = []
    #
    # # 查找为loop的、非根节点的子节点
    # for musicplaylistitem_dict in musicplaylistitem_list:
    #     if musicplaylistitem_dict['LoopCount'] == 0:
    #         # 获取非根子节点
    #         if (musicplaylistitem_dict['PlaylistItemType'] == 1):
    #             # pprint(musicplaylistitem_dict)
    #             # print()
    #             # 获取子节点的parent，即根节点
    #             root_id = musicplaylistitem_dict['parent']['id']
    #             # print(root_id)
    #             # 获取根节点的owner，即musicplaylist容器的id
    #             args = {
    #                 'waql': ' from object "%s" ' % root_id
    #             }
    #             root_list, _ = find_obj(args)
    #
    #         # 获取根节点
    #         else:
    #             # 获取根节点的owner，即musicplaylist容器的id
    #             args = {
    #                 'waql': ' from object "%s" ' % musicplaylistitem_dict['id']
    #             }
    #             root_list, _ = find_obj(args)
    #             # pprint(root_list)
    #         for root_dict in root_list:
    #             # 添加非重复元素，set为去重，然后再转为list
    #             # 获取的是为loop的非根子节点的父级
    #             loop_musicplaylist_container_id_list = list(
    #                 set(loop_musicplaylist_container_id_list + [root_dict['owner']['id']]))

    # 还有另一种，在根节点下为loop的情况

    """对象创建测试"""
    # args = {
    #     # 选择父级
    #     "parent": "\\States\\Default Work Unit",
    #     # 创建类型名称
    #     "type": "StateGroup",
    #     "name": "aaaa",
    #     "notes": "哈哈哈哈",
    # }
    # state_group_object = client.call("ak.wwise.core.object.create", args)
    # print(state_group_object)
    # # {'id': '{74939776-1E08-438A-B54C-23F902054EFA}', 'name': 'aaaa'}

    """查找引用State的Event"""
    # args = {
    #     'waql': '"%s" select referencesTo' % (
    #         '{F8D43688-5767-4CFE-9FA9-06042DA5D7FD}')
    # }
    # refer_list, refer_id = find_obj(args)
    # # pprint(refer_list)
    # # [{'id': '{23B0CE4A-3440-40BE-A83C-5571900F523A}',
    # #   'isIncluded': True,
    # #   'name': 'Amb_Global',
    # #   'notes': '',
    # #   'parent': {'id': '{446650AE-60B4-40A0-8B14-3470DBBB83B0}', 'name': 'Amb'}},
    # #  {'id': '{1750C88C-AA43-40EC-A743-20781136AF55}',
    # #   'isIncluded': False,
    # #   'name': '',
    # #   'notes': '',
    # #   'parent': {'id': '{264BD5CA-8766-4DAE-940B-394383605B01}',
    # #              'name': 'AKE_Set_State_Gameplay_Explore'}},
    # #  {'id': '{FE343C56-AABE-4FFF-8000-DDEE6740E38F}',
    # #   'isIncluded': True,
    # #   'name': 'Set_Demo',
    # #   'notes': '',
    # #   'parent': {'id': '{2EFBFD40-ACE8-4E43-A8F4-24FAFA31A30E}', 'name': 'Set'}}]
    # for refer_dict in refer_list:
    #     if 'parent' in refer_dict:
    #         if 'AKE_Set_' in refer_dict['parent']['name']:
    #             print(refer_dict['parent'])

    """事件创建测试"""
    # args = {
    #     # 选择父级
    #     "parent": "{4466FD27-21D5-44BB-A0F1-AF801D870945}",
    #     # 创建类型名称
    #     "type": "Event",
    #     "name": "new",
    #     "notes": "aaaa",
    #     "children": [
    #         {
    #             "name": "",
    #             "type": "Action",
    #             "@ActionType": 22,
    #             "@Target": "{AFD0E1E3-FCCA-48DB-BB40-55618E1B78B4}"
    #         }]
    # }
    # aa = client.call("ak.wwise.core.object.create", args)
    # print(aa)

    """WorkUnit删除测试"""

    #
    # def delete_obj(obj_id):
    #     args = {
    #         "object": "%s" % obj_id
    #     }
    #     client.call("ak.wwise.core.object.delete", args)
    #
    #
    # delete_obj("{4BA35CAF-0CA1-4423-9552-44C46099DAEA}")

    """音乐的创建：MusicPlaylistContainer和MusicSegment创建"""
    # args = {
    #     "objects": [
    #         {
    #             "object": "\\Interactive Music Hierarchy\\Default Work Unit",
    #             "children": [
    #                 {
    #                     "type": "MusicPlaylistContainer",
    #                     "name": "MySequence",
    #                     "children": [
    #                         {
    #                             "type": "MusicSegment",
    #                             "name": "Segment1",
    #                         }
    #                     ],
    #
    #                 }
    #             ]
    #         }
    #     ],
    #     "onNameConflict": "merge",
    #     "listMode": "replaceAll"
    # }
    # obj = client.call("ak.wwise.core.object.set", args)
    # pprint(obj)

    """音乐的wav导入"""
    # args_import = {
    #     # createNew
    #     # useExisting：会增加一个新媒体文件但旧的不会删除
    #     # replaceExisting:会销毁Sound，上面做的设置都无了
    #     "importOperation": "replaceExisting",
    #     "imports": [
    #         {
    #             "audioFile": "C:\\Users\\happyelements\\Desktop\\Old\\Char_Skill_C1502_Atk2.wav",
    #             "objectPath": "\\Interactive Music Hierarchy\\Default Work Unit\\MySequence\\Segment1",
    #             # "originalsSubFolder": system_name
    #         }
    #     ]
    # }
    #
    # client.call("ak.wwise.core.audio.import", args_import)

    """获取Music的Sub Track测试"""
    # # 注意，这里需要选择MusicTrack的父级，也就是Segment，然后获取的内容为Sequences
    # args = {
    #     'waql': '"%s" select descendants' % (
    #         '{CF58A7CB-4A59-49EA-8735-4F6FB6817F44}')
    # }
    # music_track, _ = find_obj(args)
    # pprint(music_track)

    """Music的Stinger创建测试"""
    # 仅23版可使用
    # args = {
    #     "objects": [
    #         {
    #             "object": "{36137DAE-C9CC-4025-BE36-C5DCA3DD80B5}",
    #             "@Trigger": "\\Triggers\\Default Work Unit\\Stin_Map_A02_Combat_Boss_Trans_Piano"
    #         }
    #     ]
    # }
    # client.call("ak.wwise.core.object.set", args)
    #
    # # args = {
    # #     "object": "{5691B17A-B70A-4575-AA75-6265F9685C5F}",
    # #     "reference": "\\Triggers\\Default Work Unit\\Stin_Map_A02_Combat_Boss_Trans_Piano",
    # #     "value": "Trigger"
    # # }
    # # client.call("ak.wwise.core.object.setReference", args)
    #
    # # Stinger查找
    # args = {
    #     'waql': 'from type MusicStinger'
    # }
    # music_track, _ = find_obj(args)
    # pprint(music_track)

    """调用头文件的waql_from_type获取"""
    # obj_list = client.call("ak.wwise.core.object.get",
    #                        waapi_h.waql_from_type("Event"),
    #                        options=config.options)['return']
    # musicplaylistitem_list, _, _ = waapi_h.find_obj(obj_list)
    # pprint(musicplaylistitem_list)

    """查找Amb"""
    # args = {
    #     'waql': ' "{23B0CE4A-3440-40BE-A83C-5571900F523A}" select descendants where type = "RandomSequenceContainer" '}
    # obj_sub_list, obj_sub_id = find_obj(args)
    # extract_key = lambda d: d['id']
    # amb_2d_list = list(map(extract_key, obj_sub_list))
    # pprint(amb_2d_list)

    """RTPC创建测试"""

    # args = {
    #     "objects": [
    #         {
    #             "object": "{02AA9D36-C4CB-45D4-B0EC-9B17749258E3}",
    #             "@RTPC": [
    #                 {
    #                     "type": "RTPC",
    #                     "name": "",
    #                     "@Curve": {
    #                         "type": "Curve",
    #                         "points": [
    #                             {
    #                                 "x": -48,
    #                                 "y": 0,
    #                                 "shape": "Linear"
    #                             },
    #                             {
    #                                 "x": 0,
    #                                 "y": -6,
    #                                 "shape": "Linear"
    #                             }
    #                         ]
    #                     },
    #                     # "notes": "a new rtpc",
    #                     "@PropertyName": "OutputBusVolume",
    #                     "@ControlInput": "{5D509B17-5562-4145-83B0-414A6107374F}"
    #                     # ControlInput填的是rtpc的名子
    #                 }
    #             ]
    #         }
    #
    #     ]
    # }
    #
    # wwise_info_list = client.call("ak.wwise.core.object.set", args)
    # pprint(wwise_info_list)
    # RTPC的类型就是RTPC，id是@RTPC的id
    # {'objects': [{'@RTPC': [{'id': '{368C32DE-C9E1-4950-ABA3-78E33FF85CAF}',
    #                          'name': ''}],
    #               'id': '{02AA9D36-C4CB-45D4-B0EC-9B17749258E3}',
    #               'name': 'Mon'}]}

    """获取选中对象的属性"""
    """获取选中对象的打包流媒体的路径"""
    # objects_list = client.call("ak.wwise.ui.getSelectedObjects")['objects']
    # obj_id = objects_list[0]['id']
    # print(obj_id)
    # args = {
    #     'waql': '"%s"' % obj_id
    # }
    # options = {
    #     'return': ['name', 'Color', 'maxDurationSource', 'type', 'pluginname', 'Effect']
    # }
    # # 存储了所有Event的单元结构
    # wwise_info_list = client.call("ak.wwise.core.object.get", args, options=options)['return']
    # pprint(wwise_info_list)

    """获取某个id的obj的信息"""
    # args = {
    #     'waql': '"%s" select Effects where notes="%s"' % ("{02AA9D36-C4CB-45D4-B0EC-9B17749258E3}", "a new rtpc")
    # }
    # options = {
    #     'return': ['name', 'Color', 'maxDurationSource', 'type', 'pluginname', 'parent', 'owner', "RTPC", 'notes', 'id',
    #                'ControlInput']
    # }
    # # 存储了所有Event的单元结构
    # wwise_info_list = client.call("ak.wwise.core.object.get", args, options=options)['return']
    # pprint(wwise_info_list)

    """为obj添加相应effects:报错"""
    # args = {
    #     "objects": [
    #         {
    #             "object": "{8E4DD34E-5139-4EF2-9494-5303775493DE}",
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
    # result = client.call("ak.wwise.core.object.set", args)

    """效果器查找"""
    # args = {
    #     'waql': '"%s"  ' % "{2D86BD0B-BC69-4416-8368-FED6C2E956B7}"
    # }
    # options = {
    #     'return': ['name', 'Color', 'maxDurationSource', 'type', 'pluginname', 'parent', 'owner', "RTPC", 'notes', 'id',
    #                'ControlInput']
    # }
    # # 存储了所有Event的单元结构
    # wwise_info_list = client.call("ak.wwise.core.object.get", args, options=options)['return']
    # pprint(wwise_info_list)

    """效果器设置"""
    # 需要先超越父级
    args = {
        "objects": [
            {
                "object": "{372CAE82-5A26-4877-B844-D6F780045924}",
                "@Effect0": "{BCB310CA-FC40-4CB6-A454-F711631441DA}"
            }
        ]
    }
    client.call("ak.wwise.core.object.set", args)

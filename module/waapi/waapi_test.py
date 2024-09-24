from waapi import WaapiClient, CannotConnectToWaapiException
from pprint import pprint

with WaapiClient() as client:
    def find_obj(args):
        options = {
            'return': ['name', 'id', 'notes', 'originalWavFilePath', 'isIncluded', 'IsLoopingEnabled',
                       'musicPlaylistRoot', 'LoopCount', 'PlaylistItemType', 'owner', 'parent']

        }
        obj_sub_list = client.call("ak.wwise.core.object.get", args, options=options)['return']
        if not obj_sub_list:
            obj_sub_id = ""
        else:
            obj_sub_id = obj_sub_list[0]['id']
        return obj_sub_list, obj_sub_id


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
    #     if (musicplaylistitem_dict['PlaylistItemType'] == 1) and (musicplaylistitem_dict['LoopCount'] == 0):
    #         # pprint(musicplaylistitem_dict)
    #         # print()
    #         # 获取子节点的parent，即根节点
    #         root_id = musicplaylistitem_dict['parent']['id']
    #         # print(root_id)
    #         # 获取根节点的owner，即musicplaylist容器的id
    #         args = {
    #             'waql': ' from object "%s" ' % root_id
    #         }
    #         root_list, _ = find_obj(args)
    #         # pprint(root_list)
    #         for root_dict in root_list:
    #             # 添加非重复元素，set为去重，然后再转为list
    #             loop_musicplaylist_container_id_list = list(set(loop_musicplaylist_container_id_list +
    #                                                             [root_dict['owner']['id']]))
    #     # 还有另一种，在根节点下为loop的情况
    # pprint(loop_musicplaylist_container_id_list)

    """对象创建测试"""
    args = {
        # 选择父级
        "parent": "\\States\\Default Work Unit",
        # 创建类型名称
        "type": "StateGroup",
        "name": "aaaa",
        "notes": "哈哈哈哈",
    }
    state_group_object = client.call("ak.wwise.core.object.create", args)
    print(state_group_object)
    # {'id': '{74939776-1E08-438A-B54C-23F902054EFA}', 'name': 'aaaa'}


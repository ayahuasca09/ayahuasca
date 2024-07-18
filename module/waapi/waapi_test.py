from waapi import WaapiClient, CannotConnectToWaapiException
from pprint import pprint

with WaapiClient() as client:
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

    # """对象删除测试"""
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
    args = {
        "sources": [
            {
                "input": "F:\\pppppy\\SP\\module\\waapi\\waapi_Auto_Import_ExternalSource\\ExternalSource.wsources",
                "platform": "Windows",
                "output": 'S:\\Ver_1.0.0\\Project\\Content\\Audio\\GeneratedExternalSources\\Windows'
            },
            {
                "input": "F:\\pppppy\\SP\\module\\waapi\\waapi_Auto_Import_ExternalSource\\ExternalSource.wsources",
                "platform": "Android",
                "output": 'S:\\Ver_1.0.0\\Project\\Content\\Audio\\GeneratedExternalSources\\Android'
            },
            {
                "input": "F:\\pppppy\\SP\\module\\waapi\\waapi_Auto_Import_ExternalSource\\ExternalSource.wsources",
                "platform": "iOS",
                "output": 'S:\\Ver_1.0.0\\Project\\Content\\Audio\\GeneratedExternalSources\\iOS'
            }
        ]
    }

    gen_log = client.call("ak.wwise.core.soundbank.convertExternalSources", args)

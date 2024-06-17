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
    args_import = {
        # createNew
        # useExisting：会增加一个新媒体文件但旧的不会删除
        # replaceExisting:会销毁Sound，上面做的设置都无了
        "importOperation": "replaceExisting",
        "default": {
            "importLanguage": "SFX"
        },
        "imports": [
            {
                "audioFile": r"F:\pppppy\SP\module\waapi\waapi_音频资源导入自动化\Media_Temp.wav",
                "objectPath": '\\Actor-Mixer Hierarchy\\v1\\Char\\Char\\Char_Skill\\Char_Skill\\Char_Skill_C01\\Char_Skill_C01\\Char_Skill_C01_Strafe_Enter\\<Sound SFX>My SFX 0',
                "originalsSubFolder": "Char"
                #                                                         名为Test 0的顺序容器            名为My SFX 0 的音效
                # "objectPath": "\\Actor-Mixer Hierarchy\\Default Work Unit\\<Sequence Container>Test 0\\<Sound SFX>My SFX 0"
            }
        ]
    }
    # 定义返回结果参数，让其只返回 Windows 平台下的信息，信息中包含 GUID 和新创建的对象名
    opts = {
        "platform": "Windows",
        "return": [
            "id", "name"
        ]
    }
    client.call("ak.wwise.core.audio.import", args_import, options=opts)

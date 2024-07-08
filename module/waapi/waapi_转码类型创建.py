from waapi import WaapiClient, CannotConnectToWaapiException
from pprint import pprint


# 创建Conversion
# 设置Channel
# 但是无法设置Format，只能手动设置

# 创建转码设置
def waapi_create_conversion(name):
    args = {
        "parent": "\\Conversion Settings\\User Conversion Settings",
        "type": "Conversion",
        "name": name,
        # 若已存在则会返回WampClientAutobahn (ERROR): ('ApplicationError(error=<ak.wwise.invalid_arguments>, args=[], '
        #  'kwargs={\'message\': "\'parent\' already has child with same name.", '
        #  "'details': {'procedureUri': 'ak.wwise.core.object.create'}}, "
        #  'enc_algo=None, callee=None, callee_authid=None, callee_authrole=None, '
        #  'forward_for=None)')
        "onNameConflict": "fail",
    }
    return args


with WaapiClient() as client:
    # 为不同conversion设置不同channels
    def waapi_conversion_settings(name, Android_Channel, iOS_Channel, Win_Channel):
        args = waapi_create_conversion(name)
        guid = client.call("ak.wwise.core.object.create", args)['id']
        # print(guid)
        convert_Settings_Android = {
            "object": guid,
            "platform": "Android",
            "property": "Channels",
            "value": Android_Channel,
        }
        client.call("ak.wwise.core.object.setProperty", convert_Settings_Android)

        convert_Settings_iOS = {
            "object": guid,
            "platform": "iOS",
            "property": "Channels",
            "value": iOS_Channel,
        }
        client.call("ak.wwise.core.object.setProperty", convert_Settings_iOS)

        convert_Settings_Win = {
            "object": guid,
            "platform": "Windows",
            "property": "Channels",
            "value": Win_Channel,
        }
        client.call("ak.wwise.core.object.setProperty", convert_Settings_Win)


    # 自动创建转码设置
    # 1.创建自定义转码unit
    args = {
        "parent": "\\Conversion Settings",
        "type": "WorkUnit",
        "name": "User Conversion Settings",
        # 若已存在则会返回WampClientAutobahn (ERROR): ('ApplicationError(error=<ak.wwise.invalid_arguments>, args=[], '
        #  'kwargs={\'message\': "\'parent\' already has child with same name.", '
        #  "'details': {'procedureUri': 'ak.wwise.core.object.create'}}, "
        #  'enc_algo=None, callee=None, callee_authid=None, callee_authrole=None, '
        #  'forward_for=None)')
        "onNameConflict": "fail"

    }
    user_convert_unit = client.call("ak.wwise.core.object.create", args)
    # print(user_convert_unit)

    # 2.创建转码设置
    # 0	Mono
    # 1	Stereo
    # 2	Mono Drop
    # 3	Stereo Drop
    # 4	As Input

    # Music
    waapi_conversion_settings("Music", 4, 4, 4)

    # VO
    waapi_conversion_settings("VO", 0, 0, 0)

    # SFX_UI_Stereo
    waapi_conversion_settings("SFX_UI_Stereo", 4, 4, 4)

    # SFX_Long_Mono
    waapi_conversion_settings("SFX_Long_Mono", 0, 0, 0)

    # SFX_Long_Stereo
    waapi_conversion_settings("SFX_Long_Stereo", 4, 4, 4)

    # SFX_Short_Mono
    waapi_conversion_settings("SFX_Short_Mono", 0, 0, 0)

    # SFX_Short_Stereo
    waapi_conversion_settings("SFX_Short_Stereo", 4, 4, 4)

    # SFX_AMB_LowQuality
    waapi_conversion_settings("SFX_AMB_LowQuality", 4, 4, 4)

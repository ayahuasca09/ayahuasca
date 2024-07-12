import openpyxl
import json
import sys
from os.path import abspath, dirname
import os
from pprint import pprint
import re
from waapi import WaapiClient
import shutil
from openpyxl.cell import MergedCell

# 文件所在目录
py_path = ""
if hasattr(sys, 'frozen'):
    py_path = dirname(sys.executable)
elif __file__:
    py_path = dirname(abspath(__file__))

"""获取json文件路径"""
json_name = '导入规则.json'
json_path = os.path.join(py_path, json_name)

# 打开json文件并转为python字典
with open(json_path, 'r', encoding='UTF-8') as jsfile:
    js_dict = json.load(jsfile)

# 获取Wwise的配置
wwise_dict = js_dict["Wwise"]
# 获取Wwise工程路径
wwise_proj_path = wwise_dict['WwiseProj_Root']
# 获取Wwise媒体资源的路径
wwise_sfx_path = os.path.join(wwise_proj_path, "Originals", "SFX")
# print(wwise_media_path)
wwise_vo_path = os.path.join(wwise_proj_path, "Originals", "Voices")
# 获取Wwise的Audio根路径
wwise_audio_root = "{5EA7BBF8-64C0-4A50-8821-A07E8BE21D68}"
# 获取Wwise的Music路径
wwise_music_root = "{49FED366-B530-42A8-993B-924873AAB707}"

# 转码设置的配置
# 配置
# 时长：当超过此时长则为长音效
duration_long = 3

# 不同转码格式的路径
default = "\\Conversion Settings\\Default Work Unit\\Default Conversion Settings"
Music = "\\Conversion Settings\\User Conversion Settings\\Music"
SFX_AMB_LowQuality = "\\Conversion Settings\\User Conversion Settings\\SFX_AMB_LowQuality"
SFX_Long_Mono = '\\Conversion Settings\\User Conversion Settings\\SFX_Long_Mono'
SFX_Long_Stereo = '\\Conversion Settings\\User Conversion Settings\\SFX_Long_Stereo'
SFX_Short_Mono = '\\Conversion Settings\\User Conversion Settings\\SFX_Short_Mono'
SFX_Short_Stereo = '\\Conversion Settings\\User Conversion Settings\\SFX_Short_Stereo'
SFX_UI_Stereo = "\\Conversion Settings\\User Conversion Settings\\SFX_UI_Stereo"
VO = "\\Conversion Settings\\User Conversion Settings\\VO"

# 获取Wwise的Sound根路径
wwise_sfx_root = "{5CB72D1B-8D4B-484D-8B9D-FC52BD496843}"

"""*****************功能检测区******************"""
"""log打印"""


def print_log(info):
    print(info)

    """清理.prof文件"""


def clear_file_prof():
    for file in os.listdir(wwise_proj_path):
        if ".prof" in file:
            # pprint(file)
            file_path = os.path.join(wwise_proj_path, file)
            os.remove(file_path)
            print_log("[文件清理]" + file + "已删除")


"""获取所有媒体文件路径并返回列表"""


def get_all_media_path(media_path):
    media_list = []
    # 遍历文件夹下的所有子文件
    # 绝对路径，子文件夹，文件名
    for root, dirs, files in os.walk(media_path):
        for file in files:
            media_list.append(os.path.join(root, file))
    return media_list


with WaapiClient() as client:
    """*****************Wwise功能区******************"""

    """设置对象引用"""


    def set_obj_refer(obj_id, refer, value):
        args = {
            "object": obj_id,
            "reference": refer,
            "value": value
        }
        client.call("ak.wwise.core.object.setReference", args)


    """设置对象属性"""


    def set_obj_property(obj_id, property, value):
        args = {
            "object": obj_id,
            "property": property,
            "value": value
        }
        client.call("ak.wwise.core.object.setProperty", args)
        # print(name_2 + ":" + str(trigger_rate))


    """查找对象"""


    def find_obj(args):
        options = {
            'return': ['name', 'id', 'path', 'notes', 'originalWavFilePath', 'maxDurationSource', '3DSpatialization']

        }
        obj_sub_list = client.call("ak.wwise.core.object.get", args, options=options)['return']
        if not obj_sub_list:
            obj_sub_id = ""
            obj_sub_path = ""
        else:
            obj_sub_id = obj_sub_list[0]['id']
            obj_sub_path = obj_sub_list[0]['path']
        return obj_sub_list, obj_sub_id, obj_sub_path


    """设置转码格式"""


    def set_conversion_type(myid, convert_type):
        # 超越父级设置
        set_obj_property(myid, "OverrideConversion", True)

        # 转码格式设置
        set_obj_refer(myid, "Conversion", convert_type)


    """清理未被引用的媒体资源"""


    def clear_file_meida_no_refer(media_list):
        # 查找所有sound
        sound_list, sound_id, _ = find_obj(
            {'waql': ' "%s" select descendants where type = "Sound" ' % wwise_audio_root})
        # pprint(sound_list)

        # 获取Wwise中有引用的文件路径
        wwise_refer_media_list = []
        for sound_dict in sound_list:
            if 'originalWavFilePath' in sound_dict:
                wwise_refer_media_list.append(sound_dict['originalWavFilePath'])
        # pprint(wwise_refer_media_list)

        for media in media_list:
            if media not in wwise_refer_media_list:
                if os.path.isfile(media):
                    # 音乐不删
                    if "Music" in media:
                        if ".wav" in media:
                            pass
                        else:
                            os.remove(media)
                            print_log("[文件清理]" + media + "已删除")
                    else:
                        os.remove(media)
                        print_log("[文件清理]" + media + "已删除")


    """不需要的文件清理"""


    def clear_file():
        # 获取媒体文件列表
        sfx_list = get_all_media_path(wwise_sfx_path)
        vo_list = get_all_media_path(wwise_sfx_path)
        # 清理未被引用的媒体资源(包括.akd和.dat)
        clear_file_meida_no_refer(sfx_list)
        clear_file_meida_no_refer(vo_list)
        # 清理.prof文件
        clear_file_prof()


    """转码自动设置"""


    def set_conversion():
        # 各类型的转码类型判断及设置
        # 语音
        if re.search(r"^(VO)", sound_dict['name']):
            set_conversion_type(sound_dict['id'], VO)
        # CG
        elif re.search(r"^(CG)", sound_dict['name']):
            set_conversion_type(sound_dict['id'], SFX_Long_Stereo)
        # Sys
        elif re.search(r"^(Sys)", sound_dict['name']):
            set_conversion_type(sound_dict['id'], SFX_UI_Stereo)
        else:
            # 其他类进行2d/3d，长/短音效检测
            if '3DSpatialization' in sound_dict.keys():
                # 获取最大时长
                sound_duration = sound_dict['maxDurationSource']['trimmedDuration']

                if sound_duration > duration_long:
                    # print('此为长声音')

                    # 判定是否为3D
                    # Value	显示名称
                    # 0	None
                    # 1	Position
                    # 2	Position + Orientation
                    if sound_dict['3DSpatialization'] == 0:
                        # print('此为2D长声音')
                        set_conversion_type(sound_dict['id'], SFX_Long_Stereo)

                    else:

                        # print('此为3D长声音')
                        set_conversion_type(sound_dict['id'], SFX_Long_Mono)

                else:
                    # 此为短声音
                    if sound_dict['3DSpatialization'] == 0:
                        # print('此为2D短声音')
                        set_conversion_type(sound_dict['id'], SFX_Short_Stereo)
                    else:
                        # print('此为3D短声音')
                        set_conversion_type(sound_dict['id'], SFX_Short_Mono)


    """针对LP的处理"""


    def set_lp():
        # 语音
        if re.search(r"(_LP)$", sound_dict['name']):
            set_obj_property(sound_dict['id'], "IsLoopingEnabled", True)
            set_obj_property(sound_dict['id'], "IsStreamingEnabled", True)
            set_obj_property(sound_dict['id'], "IsNonCachable", False)
            set_obj_property(sound_dict['id'], "IsZeroLatency", True)
            # print(sound_dict['name'])

    """*****************主程序处理******************"""
    # 不需要的文件清理
    clear_file()

    # 获取所有Sound
    sound_list, sound_id, _ = find_obj(
        {'waql': ' "%s" select descendants where type = "Sound" ' % wwise_sfx_root})
    # pprint(sound_list)

    # 仅针对Sound设置
    for sound_dict in sound_list:
        # 转码自动设置
        set_conversion()

        # 针对lp资源的设置
        set_lp()

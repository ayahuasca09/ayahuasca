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


    """*****************主程序处理******************"""
    # 不需要的文件清理
    clear_file()

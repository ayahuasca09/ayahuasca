import sys
from os.path import abspath, dirname
import os
from pprint import pprint
import re
from waapi import WaapiClient

import comlib.config as config
import comlib.audio_h as audio_h

# 文件所在目录
py_path = ""
if hasattr(sys, 'frozen'):
    py_path = dirname(sys.executable)
elif __file__:
    py_path = dirname(abspath(__file__))

# 获取Wwise工程路径
wwise_proj_path = config.wwise_proj_path
# 获取Wwise媒体资源的路径
wwise_sfx_path = os.path.normpath(config.original_path)
# print(wwise_media_path)
wwise_vo_path = os.path.normpath(config.wwise_vo_media_path)
# print(wwise_vo_path)
# 获取Wwise的Audio根路径
wwise_audio_root = "{5EA7BBF8-64C0-4A50-8821-A07E8BE21D68}"
# 获取Wwise的Music路径
wwise_music_root = config.music_path

# 转码设置的配置
# 配置
# 时长：当超过此时长则为长音效
duration_long = config.duration_long

# 不同转码格式的路径
default = "\\Conversion Settings\\Default Work Unit\\Default Conversion Settings"
Music = "\\Conversion Settings\\User Conversion Settings\\Music"
SFX_AMB_Low = "\\Conversion Settings\\User Conversion Settings\\SFX_AMB_Low"
SFX_AMB_Low_Mono = "\\Conversion Settings\\User Conversion Settings\\SFX_AMB_Low_Mono"
SFX_Long_Mono = '\\Conversion Settings\\User Conversion Settings\\SFX_Long_Mono'
SFX_Long_Stereo = '\\Conversion Settings\\User Conversion Settings\\SFX_Long_Stereo'
SFX_Short_Mono = '\\Conversion Settings\\User Conversion Settings\\SFX_Short_Mono'
SFX_Short_Stereo = '\\Conversion Settings\\User Conversion Settings\\SFX_Short_Stereo'
SFX_UI_Stereo = "\\Conversion Settings\\User Conversion Settings\\SFX_UI_Stereo"
VO = "\\Conversion Settings\\User Conversion Settings\\VO"

# 获取Wwise的Sound根路径
wwise_sfx_root = "{5CB72D1B-8D4B-484D-8B9D-FC52BD496843}"
# 获取wwisecache路径
wwise_cache_path = config.wwise_cache_path

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


    # 设置对象引用
    def set_obj_refer(obj_id, refer, value):
        args = {
            "object": obj_id,
            "reference": refer,
            "value": value
        }
        client.call("ak.wwise.core.object.setReference", args)


    """设置新的占位Sound资源超越父级为灰色"""


    def set_sound_color_default(obj_id):
        args_property = {
            "object": obj_id,
            "property": "OverrideColor",
            "value": False,
        }
        client.call("ak.wwise.core.object.setProperty", args_property)


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
            'return': ['name', 'id', 'path', 'notes', 'originalWavFilePath', 'maxDurationSource', '3DSpatialization',
                       'type', 'IsLoopingEnabled']

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


    """清理未被引用的语音资源"""


    def clear_file_vo_no_refer(vo_list):
        # 查找所有sound
        vo_sound_list, _, _ = find_obj(
            {'waql': ' "%s" select descendants where type = "Sound" and childrenCount=4' % wwise_audio_root})
        # pprint(sound_list)

        # 获取Wwise中所有语音文件的路径
        wwise_vo_path_list = []
        for vo_sound_dict in vo_sound_list:
            sound_language_list, _, _ = find_obj(
                {'waql': ' "%s" select children  ' %
                         vo_sound_dict['id']})
            if sound_language_list:
                for sound_language_dict in sound_language_list:
                    # print(sound_language_dict['originalWavFilePath'])
                    wwise_vo_path_list.append(sound_language_dict['originalWavFilePath'])

        # pprint(wwise_vo_path_list)
        # 在文件夹中查找是否有Wwise中的语音路径，若没有则为无引用
        for vo_path in vo_list:
            if vo_path not in wwise_vo_path_list:
                pass
                if os.path.isfile(vo_path):
                    os.remove(vo_path)
                    print_log("[文件清理]" + vo_path + "已删除")


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
                    if "Mus" or "Stin" in media:
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
        vo_list = get_all_media_path(wwise_vo_path)
        # 'S:\\chen.gong_DCC_Audio\\Audio\\SilverPalace_WwiseProject\\Originals\\Voices\\Korean\\VO_Game_World_C1001_06.wav',
        #  'S:\\chen.gong_DCC_Audio\\Audio\\SilverPalace_WwiseProject\\Originals\\Voices\\Korean\\VO_Game_World_C1001_07.akd',
        #  'S:\\chen.gong_DCC_Audio\\Audio\\SilverPalace_WwiseProject\\Originals\\Voices\\Korean\\VO_Game_World_C1001_07.wav',
        #  'S:\\chen.gong_DCC_Audio\\Audio\\SilverPalace_WwiseProject\\Originals\\Voices\\Korean\\VO_Game_World_C1001_08.akd',
        # 清理未被引用的媒体资源(无引用的容器和.wav以及.akd和.dat)
        clear_file_meida_no_refer(sfx_list)
        clear_file_vo_no_refer(vo_list)
        # 清理.prof文件
        clear_file_prof()


    """转码自动设置"""


    def set_conversion():
        # 各类型的转码类型判断及设置
        # 语音
        if re.search(r"^(VO)", sound_dict['name']):
            pass
            # set_conversion_type(sound_dict['id'], VO)
        # CG
        elif re.search(r"^(CG)", sound_dict['name']):
            pass
            # set_conversion_type(sound_dict['id'], SFX_Long_Stereo)
        # Sys
        elif re.search(r"^(Sys)", sound_dict['name']):
            pass
            # set_conversion_type(sound_dict['id'], SFX_UI_Stereo)
        # Amb
        elif re.search(r"^(Amb)", sound_dict['name']):
            if "Point" in sound_dict['name']:
                set_conversion_type(sound_dict['id'], SFX_AMB_Low_Mono)
            else:
                set_conversion_type(sound_dict['id'], SFX_AMB_Low)
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


    # 不管是否为循环音效，应该是长音效才开流

    def set_lp():
        # 需要去除随机后缀
        no_random_name = re.sub(r"(_R\d{2,4})$", "", sound_dict['name'])
        # 带LP后缀的处理
        if re.search(r"(_LP)$", no_random_name):
            set_obj_property(sound_dict['id'], "IsLoopingEnabled", True)


    def set_stream():
        flag = 0
        for lower_cache_path in lower_cache_list:
            if sound_dict['name'] in lower_cache_path:
                flag = 1
                set_obj_property(sound_dict['id'], "IsStreamingEnabled", False)
                break

        # 长音效也要开流
        if 'maxDurationSource' in sound_dict:
            if 'trimmedDuration' in sound_dict['maxDurationSource']:
                sound_duration = sound_dict['maxDurationSource']['trimmedDuration']
                # 长音效开流
                if sound_duration > config.duration_long:
                    flag = 0
                    # print(sound_dict['name'] + "长音效开流")

        if 'VO' in sound_dict['name']:
            flag = 2
            set_obj_property(sound_dict['id'], "IsStreamingEnabled", True)
            set_obj_property(sound_dict['id'], "IsNonCachable", True)
            set_obj_property(sound_dict['id'], "IsZeroLatency", True)

        # 大于32kb开流
        if flag == 0:
            set_obj_property(sound_dict['id'], "IsStreamingEnabled", True)
            # 环境声不需要零延迟
            if "Amb" in sound_dict['name']:
                set_obj_property(sound_dict['id'], "IsZeroLatency", False)
            else:
                set_obj_property(sound_dict['id'], "IsZeroLatency", True)
            # 循环声和语音需要禁用缓存
            if 'IsLoopingEnabled' in sound_dict:
                if sound_dict['IsLoopingEnabled']:
                    set_obj_property(sound_dict['id'], "IsNonCachable", True)
                else:
                    set_obj_property(sound_dict['id'], "IsNonCachable", False)
            else:
                if "VO" in sound_dict['name']:
                    set_obj_property(sound_dict['id'], "IsNonCachable", True)
                else:
                    set_obj_property(sound_dict['id'], "IsNonCachable", False)


    """检测是否静音：若非静音则跟随父级颜色"""


    def check_is_silence():
        if 'originalWavFilePath' in sound_dict:
            # pprint(sound_dict['originalWavFilePath'])
            if not audio_h.is_silent(sound_dict['originalWavFilePath'], silence_threshold=-50.0):
                set_sound_color_default(sound_dict['id'])
            # else:
            #     print(sound_dict['name'])


    """获取为.wav且文件大小大于32kb的文件"""


    def get_size_path_list():
        for cache_path in cache_list:
            if '.wem' in cache_path:
                if os.path.isfile(cache_path):
                    file_size = os.path.getsize(cache_path)
                    # 将字节转换为千字节（KB）
                    file_size_kb = file_size / 1024
                    if file_size_kb < config.stream_size:
                        if cache_path not in lower_cache_list:
                            lower_cache_list.append(cache_path)
                            # print(cache_path)
                            # print(file_size_kb)


    """*****************主程序处理******************"""
    # 获取cache文件路径
    cache_list = get_all_media_path(wwise_cache_path)
    # pprint(cache_list)
    # 获取制定大小之内的cache文件列表
    lower_cache_list = []
    # 获取为.wav且文件大小大于32kb的文件
    get_size_path_list()
    # pprint(lower_cache_list)

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

        # 检测是否静音：若非静音则跟随父级颜色
        check_is_silence()

        # 针对流的处理
        set_stream()

    # 针对mus/stin处理
    mus_list, _, _ = find_obj(
        {'waql': ' "%s" select descendants where type = "MusicTrack" ' % wwise_music_root})
    # pprint(mus_list)
    # 转码需要设置为Mus
    for mus_dict in mus_list:
        set_conversion_type(mus_dict['id'], Music)
        # 开流
        set_obj_property(mus_dict['id'], "IsStreamingEnabled", True)
        set_obj_property(mus_dict['id'], "IsZeroLatency", True)

    # 打包
    args = {
        "soundbanks": [
            {"name": "AKE_Play_Mus_Global"}
        ],
        "writeToDisk": True,
        # "clearAudioFileCache": True
    }

    gen_log = client.call("ak.wwise.core.soundbank.generate", args)
    pprint(gen_log)

os.system("pause")

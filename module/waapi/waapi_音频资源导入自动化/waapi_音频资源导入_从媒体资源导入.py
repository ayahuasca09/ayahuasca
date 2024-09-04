import openpyxl
import json
import sys
from os.path import abspath, dirname
import os
from pprint import pprint
import re
from waapi import WaapiClient
import shutil

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
wwise_proj_path = wwise_dict['WwiseProj_Root']
wwise_vo_media_path = wwise_proj_path + "\\Originals\\Voices"

# 是否通过
is_pass = True

"""*****************功能检测区******************"""
"""R01报错"""


def check_is_R01(media_name):
    if "_R01" in media_name:
        print_error("：_R01需要去掉,导入失败")
        no_r01_name = re.sub(r"(_R01)$", "", media_name)
        new_path = os.path.join(
            wav_path, no_r01_name + '.wav')
        os.rename(file_wav_dict[wav_name], new_path)


"""报错捕获"""


def print_warning(warn_info):
    print("[warning]" + warn_info)


def print_error(error_info):
    global is_pass
    is_pass = False
    print("[error]" + no_wav_name + error_info)

    """遍历文件目录获取文件名称和路径
    """


"""获取文件名称和路径"""


def get_type_file_name_and_path(file_type, dir_path):
    file_dict = {}
    file_list = []
    # 遍历文件夹下的所有子文件
    # 绝对路径，子文件夹，文件名
    for root, dirs, files in os.walk(dir_path):
        # {'name': ['1111.akd', 'Creature Growls 4.akd', 'Sonic Salute_005_2_20.akd'],
        #  'path': 'S:\\chen.gong_DCC_Audio\\Audio\\SilverPalace_WwiseProject\\Originals\\SFX'}
        for file in files:
            if file_type in file:
                new_dict = {}
                file_dict[file] = os.path.join(
                    root, file)
                new_dict[file] = os.path.join(
                    root, file)

                file_list.append(new_dict)

    return file_dict, file_list


# 获取媒体资源文件列表
wav_path = os.path.join(py_path, "New_Media")
_, file_wav_list = get_type_file_name_and_path('.wav', wav_path)
# pprint(file_wav_dict)
# {'Char_Mov_Gen_Boot_End_Dirt_R01.wav': 'F:\\pppppy\\SP\\module\\waapi\\waapi_音频资源导入自动化\\New_Media\\Char_Mov_Gen_Boot_End_Dirt_R01.wav',
#  'Char_Mov_Gen_Boot_End_Dirt_R02.wav': 'F:\\pppppy\\SP\\module\\waapi\\waapi_音频资源导入自动化\\New_Media\\Char_Mov_Gen_Boot_End_Dirt_R02.wav',}

with WaapiClient() as client:
    """*****************Wwise功能区******************"""

    """查找对象"""


    def find_obj(args):
        options = {
            'return': ['name', 'id', 'path', 'notes']

        }
        obj_sub_list = client.call("ak.wwise.core.object.get", args, options=options)['return']
        if not obj_sub_list:
            obj_sub_id = ""
            obj_sub_path = ""
        else:
            obj_sub_id = obj_sub_list[0]['id']
            obj_sub_path = obj_sub_list[0]['path']
        return obj_sub_list, obj_sub_id, obj_sub_path


    """获取wwise对象的最长字符串父级"""


    def find_obj_parent(obj_name, waql):
        # 所有Unit获取
        parent_list, parent_id, _ = find_obj(waql)

        # 存储符合的对象的父级，最后找出最长最符合的
        parent_len = 0
        obj_parent_id = {}
        if parent_list != None:
            for parent_dict in parent_list:
                if parent_dict['name'] in obj_name:
                    if len(parent_dict['name']) > parent_len:
                        parent_len = len(parent_dict['name'])
                        obj_parent_id = parent_dict['id']
        else:
            print_error("相应的Unit目录结构未创建！需要创建")

        return obj_parent_id


    """设置新的占位Sound资源超越父级为灰色"""


    def set_sound_color_default(obj_id):
        args_property = {
            "object": obj_id,
            "property": "OverrideColor",
            "value": False,
        }
        client.call("ak.wwise.core.object.setProperty", args_property)


    """随机容器内容取代"""


    def create_rnd_container(media_name, system_name):
        # 查找该rnd是否已存在
        flag = 0
        # 去除随机容器数字
        rnd_name = re.sub(r"(_R\d{2,4})$", "", media_name)
        rnd_path = ""
        rnd_container_list, rnd_id, _ = find_obj(
            {'waql': ' "%s" select descendants where type = "RandomSequenceContainer" ' % os.path.join(
                wwise_dict['Root'], system_name)})
        for rnd_container_dict in rnd_container_list:
            if rnd_container_dict['name'] == rnd_name:
                # pprint(rnd_name + "：RandomContainer已存在")
                flag = 1
                rnd_path = rnd_container_dict['path']
                # 在容器中替换媒体资源
                import_info = import_media_in_rnd(file_wav_dict[wav_name], media_name, rnd_path)
                if import_info != None:
                    print_warning(media_name + "：已导入随机容器" + rnd_name)
                    # 颜色修改
                    sound_container_list, sound_container_id, _ = find_obj(
                        {'waql': 'from type Sound where name = "%s"' % media_name})
                    for sound_cotainer_dict in sound_container_list:
                        set_sound_color_default(sound_cotainer_dict['id'])
        if flag == 0:
            print_error("：未找到相应的随机randomcontainer，导入失败")

        return rnd_path, rnd_name


    """媒体资源导入"""


    def import_media_in_rnd(source_path, media_name, rnd_path):
        args_import = {}
        # 导入语音
        if word_list[0] == "VO":
            if 'Chinese' in source_path:
                args_import = {
                    "importOperation": "useExisting",
                    "imports": [
                        {
                            "audioFile": source_path,
                            "objectPath": rnd_path + "\\<Sound Voice>" + media_name + "\\<AudioFileSource>" + "CN_" + media_name,
                            "importLanguage": "Chinese"
                        }
                    ]
                }
            elif 'English' in source_path:
                args_import = {
                    "importOperation": "useExisting",
                    "imports": [

                        {
                            "audioFile": source_path,
                            "objectPath": rnd_path + "\\<Sound Voice>" + media_name + "\\<AudioFileSource>" + "EN_" + media_name,
                            "importLanguage": "English"
                        }
                    ]
                }
            elif 'Japanese' in source_path:
                args_import = {
                    "importOperation": "useExisting",
                    "imports": [

                        {
                            "audioFile": source_path,
                            "objectPath": rnd_path + "\\<Sound Voice>" + media_name + "\\<AudioFileSource>" + "JP_" + media_name,
                            "importLanguage": "Japanese"
                        }
                    ]
                }
            elif 'Korean' in source_path:
                args_import = {
                    "importOperation": "useExisting",
                    "imports": [
                        {
                            "audioFile": source_path,
                            "objectPath": rnd_path + "\\<Sound Voice>" + media_name + "\\<AudioFileSource>" + "KR_" + media_name,
                            "importLanguage": "Korean"
                        }
                    ]
                }

        # 导入音效
        else:
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
                        "audioFile": source_path,
                        "objectPath": rnd_path + '\\<Sound SFX>' + media_name,
                        "originalsSubFolder": word_list[0]
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
        if args_import:
            import_info = client.call("ak.wwise.core.audio.import", args_import, options=opts)
            return import_info


    """导入语言函数"""


    def import_vo(source_path, vo_sound_dict, vo_name, lang):
        opts = {
            "platform": "Windows",
            "return": [
                "id", "name"
            ]
        }

        args_import = {
            "importOperation": "useExisting",
            "imports": [
                {
                    "audioFile": source_path,
                    "objectPath": vo_sound_dict['path'] + "\\<AudioFileSource>" + vo_name,
                    "importLanguage": lang
                }
            ]
        }
        import_info = client.call("ak.wwise.core.audio.import", args_import, options=opts)
        if import_info:
            print_warning(vo_name + ":缺失语言补全")


    """导入缺失语言的语音"""


    def import_no_language_vo(source_path, lang_cn, lang_en, lang_jp, lang_kr, vo_sound_dict):
        if not lang_cn:
            import_vo(source_path, vo_sound_dict, "CN_" + vo_sound_dict['name'], "Chinese")

        if not lang_en:
            import_vo(source_path, vo_sound_dict, "EN_" + vo_sound_dict['name'], "English")

        if not lang_jp:
            import_vo(source_path, vo_sound_dict, "JP_" + vo_sound_dict['name'], "Japanese")

        if not lang_kr:
            import_vo(source_path, vo_sound_dict, "KR_" + vo_sound_dict['name'], "Korean")


    """检查生成语音文件，若为缺资源的（例如随机容器）则补齐资源"""


    def check_and_create_vo():
        # 查找VO Unit的id
        _, vo_unit_id, _ = find_obj(
            {'waql': ' "%s" select children where type = "WorkUnit" and name="VO" ' %
                     wwise_dict['Root']})
        # pprint(vo_unit_id)

        # 查找所有SoundVoice
        if vo_unit_id:
            vo_sound_list, _, _ = find_obj(
                {'waql': ' "%s" select descendants where type = "Sound" and childrenCount<4 ' %
                         vo_unit_id})
            # pprint(vo_sound_list)
            for vo_sound_dict in vo_sound_list:
                if "External" not in vo_sound_dict['name']:
                    # 获取语言不全的列表，补齐其他语言
                    sound_language_list, _, _ = find_obj(
                        {'waql': ' "%s" select children  ' %
                                 vo_sound_dict['id']})
                    # pprint(sound_language_list)

                    # 媒体资源复制
                    # 复制到的目录
                    copy_catalog = os.path.join(py_path, "New_Media")
                    source_path = shutil.copy2(os.path.join(py_path, "Media_Temp.wav"),
                                               os.path.join(copy_catalog, vo_sound_dict['name'] + ".wav"))

                    # 查找以下语言是否有的标志
                    lang_cn = False
                    lang_en = False
                    lang_jp = False
                    lang_kr = False
                    for sound_language_dict in sound_language_list:
                        if "CN_VO" in sound_language_dict['name']:
                            lang_cn = True
                        elif "EN_VO" in sound_language_dict['name']:
                            lang_en = True
                        elif "JP_VO" in sound_language_dict['name']:
                            lang_jp = True
                        elif "KR_VO" in sound_language_dict['name']:
                            lang_kr = True

                    import_no_language_vo(source_path, lang_cn, lang_en, lang_jp, lang_kr, vo_sound_dict)


    """媒体资源导入的总流程"""


    def create_wwise_content(media_name, system_name):

        # 随机容器内容取代
        rnd_path, rnd_name = create_rnd_container(media_name, system_name)


    """*****************主程序处理******************"""
    # 撤销开始
    client.call("ak.wwise.core.undo.beginGroup")

    # 读取媒体资源文件
    for file_wav_dict in file_wav_list:
        for wav_name in file_wav_dict:
            # 去除wav的名字
            no_wav_name = re.sub(".wav", "", wav_name)
            # pprint(no_wav_name)
            word_list = no_wav_name.split("_")
            if word_list[0] in js_dict:
                check_is_R01(no_wav_name)
                if is_pass == True:
                    create_wwise_content(no_wav_name, word_list[0])

    # 检查生成语音文件，若为缺资源的（例如随机容器）则补齐资源
    check_and_create_vo()

    # 撤销结束
    client.call("ak.wwise.core.undo.endGroup", displayName="rnd创建撤销")

    # 清除复制的媒体资源
    # shutil.rmtree("New_Media")
    # os.mkdir("New_Media")

    # rnd创建撤销
    # client.call("ak.wwise.core.undo.undo")

    # 清除复制的媒体资源
    shutil.rmtree("New_Media")
    os.mkdir("New_Media")
    os.mkdir("New_Media/Chinese")
    os.mkdir("New_Media/English")
    os.mkdir("New_Media/Japanese")
    os.mkdir("New_Media/Korean")

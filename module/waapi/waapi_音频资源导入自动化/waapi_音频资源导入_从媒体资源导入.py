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


def print_error(error_info):
    global is_pass
    is_pass = False
    print(no_wav_name + error_info)

    """遍历文件目录获取文件名称和路径
    """


def get_type_file_name_and_path(file_type, dir_path):
    file_dict = {}
    # 遍历文件夹下的所有子文件
    # 绝对路径，子文件夹，文件名
    for root, dirs, files in os.walk(dir_path):
        # {'name': ['1111.akd', 'Creature Growls 4.akd', 'Sonic Salute_005_2_20.akd'],
        #  'path': 'S:\\chen.gong_DCC_Audio\\Audio\\SilverPalace_WwiseProject\\Originals\\SFX'}
        for file in files:
            if file_type in file:
                file_dict[file] = os.path.join(
                    root, file)

    return file_dict


# 获取媒体资源文件列表
wav_path = os.path.join(py_path, "New_Media")
file_wav_dict = get_type_file_name_and_path('.wav', wav_path)
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


    """创建新的随机容器"""


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
                # 在容器中创建媒体资源
                import_info = import_media_in_rnd(file_wav_dict[wav_name], media_name, rnd_path)
                if import_info != None:
                    pprint(media_name + "：已导入随机容器" + rnd_name)
                break
        # 不存在则创建
        # if flag == 0:
        #     # 查找所在路径
        #     rnd_parent_id = find_obj_parent(rnd_name,
        #                                     {
        #                                         'waql': ' "%s" select descendants,this where type = "ActorMixer" ' %
        #                                                 os.path.join(
        #                                                     wwise_dict['Root'], system_name)})
        #     if rnd_parent_id != None:
        #         # 创建的rnd属性
        #         args = {
        #             # 选择父级
        #             "parent": rnd_parent_id,
        #             # 创建类型机名称
        #             "type": "RandomSequenceContainer",
        #             "name": rnd_name,
        #             "@RandomOrSequence": 1,
        #             "@NormalOrShuffle": 0,
        #             "@RandomAvoidRepeatingCount": 3
        #         }
        #         rnd_container_object = client.call("ak.wwise.core.object.create", args)
        #         # 查找新创建的容器的路径
        #         _, _, rnd_path = find_obj(
        #             {'waql': ' "%s"  ' % rnd_container_object['id']})
        #
        #         pprint(rnd_name + "：RandomContainer创建")

        return rnd_path, rnd_name


    """媒体资源导入"""


    def import_media_in_rnd(source_path, media_name, rnd_path):
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
        import_info = client.call("ak.wwise.core.audio.import", args_import, options=opts)
        return import_info


    """创建事件的参数"""


    def get_event_args(event_parent_path, event_name, event_action, event_target):
        args_new_event = {
            # 上半部分属性中分别为 Event 创建后存放的路径、类型、名称、遇到名字冲突时的处理方法
            "parent": event_parent_path,
            "type": "Event",
            "name": event_name,
            "onNameConflict": "merge",
            "children": [
                {
                    # 为其创建播放行为，名字留空，使用 @ActionType 定义其播放行为为 Play，@Target 为被播放的声音对象
                    "name": "",
                    "type": "Action",
                    "@ActionType": event_action,
                    "@Target": event_target
                }
            ]
        }
        return args_new_event


    """事件生成"""


    def create_event(rnd_name, rnd_path):
        parent_id = find_obj_parent(rnd_name, {
            'waql': ' "%s" select descendants where type = "WorkUnit" ' % os.path.join(wwise_dict['Event_Root'],
                                                                                       word_list[0])})
        event_name = "AKE_" + "Play_" + rnd_name
        # 查找事件是否已存在
        event_list, _, _ = find_obj(
            {'waql': ' "%s" select descendants where type = "Event" ' %
                     wwise_dict['Event_Root']})
        flag = 0
        for event_dict in event_list:
            if event_dict['name'] == event_name:
                flag = 1
                break
        if flag == 0:
            event_args = get_event_args(parent_id, event_name, 1, rnd_path)
            client.call("ak.wwise.core.object.create", event_args)
            pprint(event_name + "事件创建")
            if "_LP" in rnd_name:
                event_name = "AKE_" + "Stop_" + rnd_name
                event_args = get_event_args(parent_id, event_name, 2, rnd_path)
                client.call("ak.wwise.core.object.create", event_args)
                pprint(event_name + "事件创建")


    """媒体资源导入的总流程"""


    def create_wwise_content(media_name, system_name):

        # 随机容器创建
        rnd_path, rnd_name = create_rnd_container(media_name, system_name)

        # 事件自动生成
        # create_event(rnd_name, rnd_path)


    """*****************主程序处理******************"""
    # 撤销开始
    client.call("ak.wwise.core.undo.beginGroup")

    # 读取媒体资源文件
    for wav_name in file_wav_dict:
        # 去除wav的名字
        no_wav_name = re.sub(".wav", "", wav_name)
        # pprint(no_wav_name)
        word_list = no_wav_name.split("_")
        if word_list[0] in js_dict:
            check_is_R01(no_wav_name)
            if is_pass == True:
                create_wwise_content(no_wav_name, word_list[0])

    # 撤销结束
    client.call("ak.wwise.core.undo.endGroup", displayName="rnd创建撤销")

    # 清除复制的媒体资源
    # shutil.rmtree("New_Media")
    # os.mkdir("New_Media")

    # rnd创建撤销
    # client.call("ak.wwise.core.undo.undo")

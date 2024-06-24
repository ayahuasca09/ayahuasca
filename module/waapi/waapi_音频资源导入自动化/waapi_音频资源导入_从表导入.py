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

"""获取.xlsx文件"""

file_names = []
for i in os.walk(py_path):
    file_names.append(i)
# pprint("输出文件夹下的文件名：")
file_name_list = file_names[0][2]

# 检测是否通过
is_pass = True
# 一级系统名称
system_name = ""

"""*****************功能检测区******************"""
"""检查描述/状态/修饰词是否在末尾"""
"""报错捕获"""


def print_error(error_info):
    global is_pass
    is_pass = False
    print(cell_sound.value + error_info)


"""检测字符串是否含有中文"""


def check_is_chinese(string):
    for ch in string:
        if u'\u4e00' <= ch <= u'\u9fff':
            return True
    return False


"""通过正则表达式检查"""


def check_by_re(pattern, name):
    pattern = str(pattern)
    name = str(name)
    result = re.search(pattern, name)
    new_name = ""
    if result == None:
        print_error("：请检查" + pattern + "是否拼写错误或未添加到列表中")
    else:
        new_name = name.replace(result.group(), "")
    return new_name


"""字符串长度检查"""


def check_by_str_length(str1, length):
    if len(str1) > length:
        if "_" in str1:
            print_error("：描述后缀名过长，限制字符数为" + str(length))
        else:
            print_error("：通过_切分的某个单词过长，每个单词长度不应超过10个字符，需要再拆分")


"""模块类型输出报错"""


def print_error_by_module(module_dict):
    module_type = ""
    for module in module_dict:
        # 如果不是最后一位，需要加|
        if module != list(module_dict.keys())[-1]:
            module_type += module + "|"
        else:
            module_type += module
    print_error("：" + module_type + "有误，请检查是否添加模块名称或是否拼写有误")


"""Amb类型检查"""


def check_by_Amb(name):
    global is_pass
    is_pass = False


"""CG类型检查"""


def check_by_CG(name):
    global is_pass
    is_pass = False


"""Imp类型检查"""


def check_by_Imp(name):
    global is_pass
    is_pass = False


"""Mus类型检查"""


def check_by_Mus(name):
    global is_pass
    is_pass = False


"""VO类型检查"""


def check_by_VO(name):
    global is_pass
    is_pass = False


"""Char类型检查"""


def check_by_Char(name):
    module_dict = js_dict[system_name]['module']
    flag = 0
    is_mov = False
    for module in module_dict:
        # 模块层查询
        if module + "_" in name:
            name = name.replace(module + "_", "")
            # 角色名层查询
            name = check_by_re(js_dict[system_name]['name'] + "_", name)
            if name != None:
                if "Mov_" in name:
                    is_mov = True
                # property
                name = check_by_re(module_dict[module]['property'] + "_*", name)
                # 长度限制查询
                if name != None:
                    check_by_str_length(name, js_dict[system_name]['length'])
                # Mov层加查询
                if is_mov == True:
                    name = check_by_re(module_dict[module]['property2'] + "_*", name)
                    if name != None:
                        name = check_by_re(module_dict[module]['property3'] + "_*", name)

            flag = 1
            break

    if flag == 0:
        print_error_by_module(module_dict)


"""Mon类型检查"""


def check_by_Mon(name):
    type_dict = js_dict[system_name]
    flag = 0
    for type in type_dict:
        if type in name:
            name = name.replace(type + "_", "")
            flag = 1
            module_dict = js_dict[system_name][type]['module']
            flag = 0
            for module in module_dict:
                # 模块层查询
                if module + "_" in name:
                    name = name.replace(module + "_", "")
                    # Boss名层查询
                    name = check_by_re(js_dict[system_name][type]['name'] + "_", name)
                    if name != None:
                        # 技能层查询
                        name = check_by_re(module_dict[module]['property'] + "_*", name)
                        # 长度限制查询
                        if name != None:
                            check_by_str_length(name, js_dict[system_name][type]['length'])

                    flag = 1
                    break

            if flag == 0:
                print_error_by_module(module_dict)
            break
    if flag == 0:
        print_error_by_module(type_dict)


def check_by_Sys(name):
    module_dict = js_dict[system_name]['module']
    flag = 0
    for module in module_dict:
        # 模块层查询
        if module + "_" in name:
            name = name.replace(module + "_", "")
            # Show类查询
            if module == "Show":
                if name != None:
                    # 角色名称
                    name = check_by_re(module_dict[module]['property'] + "_", name)
                    if name != None:
                        # 皮肤名称
                        name = check_by_re(module_dict[module]['property2'] + "_", name)
                        if name != None:
                            # 动作名称
                            name = check_by_re(module_dict[module]['property3'] + "_*", name)
                            if name != None:
                                # 长度限制查询
                                check_by_str_length(name, module_dict[module]['length'])
            flag = 1
            break
    if flag == 0:
        print_error_by_module(module_dict)


"""检查LP是否在末尾"""


def check_LP_in_last():
    if "LP" in name:
        # print("LP")
        # .*代表所有字符 _LP$代表以_LP结尾
        pattern = r".*_LP$"
        result = re.match(pattern, name)
        if result != None:
            new_name = re.sub(r"_LP$", "", name)
            return new_name
        else:
            print_error("：_LP应放在最末尾")
    else:
        return name


"""检查是否使用通用词汇"""


def check_by_com_word(name):
    com_word_dict = js_dict['Gen_Word']
    for key in com_word_dict:
        for value in com_word_dict[key]:
            if value in name:
                print_error(value + "应改为通用词" + key)
                if "Medium" in name:
                    print_error("PS：Mid表体积/重量，Med表距离")


"""获取一级系统名走不同的检测方式"""


def check_first_system_name(name):
    global system_name
    system_name = word_list[0]
    if system_name in js_dict:
        func_name = "check_by_" + system_name
        eval(func_name)(name.replace(system_name + "_", ""))
        if is_pass == True:
            create_wwise_content(cell_sound.value, system_name)
    else:
        print_error_by_module(js_dict)


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
                # pprint(rnd_name + "：RandomContainer已存在，将不再导入")
                flag = 1
                rnd_path = rnd_container_dict['path']
                break
        # 不存在则创建
        if flag == 0:
            # 查找所在路径
            rnd_parent_id = find_obj_parent(rnd_name,
                                            {
                                                'waql': ' "%s" select descendants,this where type = "ActorMixer" ' %
                                                        os.path.join(
                                                            wwise_dict['Root'], system_name)})
            if rnd_parent_id != None:
                # 创建的rnd属性
                args = {
                    # 选择父级
                    "parent": rnd_parent_id,
                    # 创建类型机名称
                    "type": "RandomSequenceContainer",
                    "name": rnd_name,
                    "@RandomOrSequence": 1,
                    "@NormalOrShuffle": 0,
                    "@RandomAvoidRepeatingCount": 3
                }
                rnd_container_object = client.call("ak.wwise.core.object.create", args)
                # 查找新创建的容器的路径
                _, _, rnd_path = find_obj(
                    {'waql': ' "%s"  ' % rnd_container_object['id']})

                # 媒体资源复制
                # 复制到的目录
                copy_catalog = os.path.join(py_path, "New_Media")
                source_path = shutil.copy2(os.path.join(py_path, "Media_Temp.wav"),
                                           os.path.join(copy_catalog, media_name + ".wav"))

                # 在容器中创建媒体资源
                import_media_in_rnd(source_path, media_name, rnd_path)

                pprint(rnd_name + "：RandomContainer创建及媒体资源导入")

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
                    "originalsSubFolder": system_name
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
                                                                                       system_name)})
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
        create_event(rnd_name, rnd_path)


    def delete_wwise_content():
        for i in file_name_list:
            if ".xlsx" in i:
                # 拼接xlsx的路径
                file_path_xlsx = os.path.join(py_path, i)
                # 获取xlsx的workbook
                wb = openpyxl.load_workbook(file_path_xlsx)
                # 获取xlsx的所有sheet
                sheet_names = wb.sheetnames
                # 加载所有工作表
                for sheet_name in sheet_names:
                    sheet = wb[sheet_name]
                    # 获取工作表第一行数据
                    for cell in list(sheet.rows)[0]:
                        if 'Sample Name' in str(cell.value):
                            # 获取音效名下的内容
                            for cell_sound in list(sheet.columns)[cell.column - 1]:
                                # 空格和中文不检测
                                if cell_sound.value != None:
                                    if check_is_chinese(cell_sound.value) == False:
                                        """❤❤❤❤数据获取❤❤❤❤"""
                                        """测试名称"""
                                        name = cell_sound.value
                                        if rnd_container['name'] == name:
                                            return
        # 随机容器引用的资产删除
        media_list, media_id, _ = find_obj(
            {'waql': ' "%s" select children where type = "Sound" ' % rnd_container['id']})
        for media_dict in media_list:
            os.remove(media_dict['originalWavFilePath'])
            pprint(media_dict['originalWavFilePath'] + "已删除")

        # 随机容器删除
        args = {
            "object": "%s" % rnd_container['id']
        }
        client.call("ak.wwise.core.object.delete", args)
        pprint(rnd_container['name'] + ":RandomContainer已删除")

        # 事件删除
        for event_dict in event_list:
            if rnd_container['name'] in event_dict['name']:
                args = {
                    "object": "%s" % event_dict['id']
                }
                client.call("ak.wwise.core.object.delete", args)
                pprint(event_dict['name'] + ":Event已删除")


    """*****************主程序处理******************"""
    # 撤销开始
    client.call("ak.wwise.core.undo.beginGroup")

    # 提取规则：只提取xlsx文件
    for i in file_name_list:
        if ".xlsx" in i:
            # 拼接xlsx的路径
            file_path_xlsx = os.path.join(py_path, i)
            # 获取xlsx的workbook
            wb = openpyxl.load_workbook(file_path_xlsx)
            # 获取xlsx的所有sheet
            sheet_names = wb.sheetnames
            # 加载所有工作表
            for sheet_name in sheet_names:
                sheet = wb[sheet_name]
                # 获取工作表第一行数据
                for cell in list(sheet.rows)[0]:
                    if 'Sample Name' in str(cell.value):
                        # 获取音效名下的内容
                        for cell_sound in list(sheet.columns)[cell.column - 1]:
                            # 空格和中文不检测
                            if cell_sound.value != None:
                                if check_is_chinese(cell_sound.value) == False:
                                    """❤❤❤❤数据获取❤❤❤❤"""
                                    """测试名称"""
                                    name = cell_sound.value
                                    # pprint(cell_sound.value)
                                    word_list = name.split("_")
                                    is_title = True
                                    for word in word_list:
                                        """判定_的每个都不能超过10个"""
                                        check_by_str_length(word, 10)
                                        """判定_的每个开头必须大写"""
                                        if len(word) > 0:
                                            # 只需要检查第一个字符，因为istitle会导致检查字符串只能首字母大写
                                            if word[0].istitle() == False and word[0].isdigit() == False:
                                                is_title = False
                                                break
                                    if is_title == False:
                                        print_error("：通过”_“分隔的每个单词开头都需要大写")

                                    # 检查LP是否在末尾
                                    name = check_LP_in_last()
                                    # print(name)

                                    # 通用词汇检查
                                    check_by_com_word(name)

                                    # 检查一级系统并索引到不同的检测方式
                                    check_first_system_name(name)
                                    # print(name)

    """删除表中没有的内容"""
    # 查找所有Rnd和Event
    rnd_container_list, rnd_id, _ = find_obj(
        {'waql': ' "%s" select descendants where type = "RandomSequenceContainer" ' % wwise_dict['Root']})
    # [{'id': '{C1BFDDC1-CA6F-45BC-8B43-15AE866AA20A}',
    #   'name': 'Char_Skill_C01_Atk1',
    #   'notes': '',
    #   'path': '\\Actor-Mixer '
    #           'Hierarchy\\v1\\Char\\Char\\Char_Skill\\Char_Skill\\Char_Skill_C01\\Char_Skill_C01\\Char_Skill_C01_Atk1'}

    event_list, event_id, _ = find_obj(
        {'waql': ' "%s" select descendants where type = "Event" ' % wwise_dict['Event_Root']})

    for rnd_container in rnd_container_list:
        delete_wwise_content()

    # 撤销结束
    client.call("ak.wwise.core.undo.endGroup", displayName="rnd创建撤销")

    # 清除复制的媒体资源
    shutil.rmtree("New_Media")
    os.mkdir("New_Media")

    # rnd创建撤销
    # client.call("ak.wwise.core.undo.undo")




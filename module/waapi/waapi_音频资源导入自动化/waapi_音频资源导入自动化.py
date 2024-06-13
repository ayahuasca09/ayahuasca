import openpyxl
import json
import sys
from os.path import abspath, dirname
import os
from pprint import pprint
import re
from waapi import WaapiClient, CannotConnectToWaapiException

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

"""*****************功能检测区******************"""
"""检查描述/状态/修饰词是否在末尾"""

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
    is_pass = False
    if result == None:
        print((cell_sound.value + "：请检查" + pattern + "是否拼写错误或未添加到列表中"))
    else:
        new_name = name.replace(result.group(), "")
        is_pass = True
    return new_name, is_pass


"""字符串长度检查"""


def check_by_str_length(str1, length):
    is_pass = True
    if len(str1) > length:
        is_pass = False
        if "_" in str1:
            print(name + "：描述后缀名过长，限制字符数为" + str(length))
        else:
            print(name + "：通过_切分的某个单词过长，每个单词长度不应超过10个字符，需要再拆分")
    return is_pass


# 已废弃按照json配置检查
def check_by_json_by_re(system_dict):
    pattern = (
        # 一级结构
            "(?P<the1>" + "^" +
            word_list[0] +
            ")" + "_" +

            # module
            "(?P<the2>" +
            system_dict['module'] +
            ")" + "_" +

            # name
            "(?P<the3>" +
            system_dict['name'] +
            ")" + "_"

            # property
                  "(?P<the4>" +
            system_dict['property'] +
            ")" + "_*"

            # 其他
                  "(?P<the5>" +
            ".*" +
            ")"
    )
    result = re.search(pattern, name)
    if result == None:
        print((name + "：中间有字段匹配错误，请检查是否为结构顺序或拼写错误"))
    else:
        data_list = re.finditer(pattern, name)
        for item in data_list:
            item_dict = item.groupdict()
            # pprint(item_dict)
            # {'the1': 'Char',
            #  'the2': 'Skill',
            #  'the3': 'C01',
            #  'the4': 'Focus',
            #  'the5': 'Ready2'}

            # 字符串长度检查
            check_by_str_length(item_dict['the5'], system_dict['length'])


"""Char类型检查"""


def check_by_char(name):
    module_dict = js_dict['Char']['module']
    flag = 0
    is_mov = False
    is_pass = 0
    for module in module_dict:
        # 模块层查询
        if module + "_" in name:
            name = name.replace(module + "_", "")
            # 角色名层查询
            name = check_by_re(js_dict['Char']['name'] + "_", name)
            if name != None:
                if "Mov_" in name:
                    is_mov = True
                # property
                name = check_by_re(module_dict[module]['property'] + "_*", name)
                # 长度限制查询
                if name != None:
                    is_pass = check_by_str_length(name, js_dict['Char']['length'])
                # Mov层加查询
                if is_mov == True:
                    name = check_by_re(module_dict[module]['property2'] + "_*", name)
                    if name != None:
                        name, is_pass = check_by_re(module_dict[module]['property3'] + "_*", name)

            flag = 1
            break

    if flag == 0:
        print(cell_sound.value + "：Char的模块（如Skill，Foley等）有误，请检查是否添加模块名称或是否拼写有误")
    return is_pass


"""Mon类型检查"""


def check_by_mon(name):
    if "Mob_" in name:
        name = name.replace("Mob_", "")
        check_by_mob(name)
    elif "Boss_" in name:
        name = name.replace("Boss_", "")
        check_by_boss(name)
    else:
        print(cell_sound.value + "：Boss、Mob拼写有误或漏写，请检查")


"""Boss类型检查"""


def check_by_boss(name):
    module_dict = js_dict['Mon']['Boss']['module']
    flag = 0
    for module in module_dict:
        # 模块层查询
        if module + "_" in name:
            name = name.replace(module + "_", "")
            # Boss名层查询
            name = check_by_re(js_dict['Mon']['Boss']['name'] + "_", name)
            if name != None:
                # 技能层查询
                name = check_by_re(module_dict[module]['property'] + "_*", name)
                # 长度限制查询
                if name != None:
                    check_by_str_length(name, js_dict['Mon']['Boss']['length'])

            flag = 1
            break

    if flag == 0:
        print(cell_sound.value + "：Boss的模块（如Body，Part等）有误，请检查是否添加模块名称或是否拼写有误")


"""Mob类型检查"""


def check_by_mob(name):
    module_dict = js_dict['Mon']['Mob']['module']
    flag = 0
    for module in module_dict:
        # 模块层查询
        if module + "_" in name:
            name = name.replace(module + "_", "")
            # 小怪名层查询
            name = check_by_re(js_dict['Mon']['Mob']['name'] + "_", name)
            if name != None:
                # 技能层查询
                name = check_by_re(module_dict[module]['property'] + "_*", name)
                # 长度限制查询
                if name != None:
                    check_by_str_length(name, js_dict['Mon']['Mob']['length'])
            flag = 1
            break

    if flag == 0:
        print(cell_sound.value + "：Mob的模块（如Skill，Foley等）有误，请检查是否添加模块名称或是否拼写有误")


def check_by_sys(name):
    module_dict = js_dict['Sys']['module']
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
        print(cell_sound.value + "：Sys的模块（如Show等）有误，请检查是否添加模块名称或是否拼写有误")


"""获取一级系统名走不同的检测方式"""


def check_first_system_name(name):
    if word_list[0] == "Amb":
        return name.replace("Amb_", "")
    elif word_list[0] == "Cg":
        return name.replace("Cg_", "")
    elif word_list[0] == "Char":
        # 如果检测通过
        if (check_by_char(name.replace("Char_", ""))) == True:
            media_import(wwise_dict['Root'], cell_sound.value)
    elif word_list[0] == "Imp":
        return name.replace("Imp_", "")
    elif word_list[0] == "Mon":
        check_by_mon(name.replace("Mon_", ""))
    elif word_list[0] == "Mus":
        return name.replace("Mus_", "")
    elif word_list[0] == "Sys":
        check_by_sys(name.replace("Sys_", ""))
    elif word_list[0] == "VO":
        return name.replace("VO_", "")
    else:
        print(name + "：一级系统名称（如Amb、Char）等有误，请检查拼写")
        return name


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
            print(cell_sound.value + "：_LP应放在最末尾")
    else:
        return name


"""检查是否使用通用词汇"""


def check_by_com_word(name):
    com_word_dict = js_dict['Gen_Word']
    for key in com_word_dict:
        for value in com_word_dict[key]:
            if value in name:
                print(cell_sound.value + "：" + value + "应改为通用词" + key)
                if "Medium" in name:
                    print("PS：Mid表体积/重量，Med表距离")


with WaapiClient() as client:
    """*****************Wwise功能区******************"""


    def media_import(path, rnd_name):
        args = {
            # 选择父级
            "parent": path,
            # 创建类型机名称
            "type": "RandomSequenceContainer",
            "name": rnd_name,
            "@RandomOrSequence": 1,
            "@NormalOrShuffle": 0,
            "@RandomAvoidRepeatingCount": 3
        }
        # 创建rnd object并获取其信息
        rnd_container_object = client.call("ak.wwise.core.object.create", args)


    """*****************主程序处理******************"""
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
                                        print(name + "：通过”_“分隔的每个单词开头都需要大写")

                                    # 检查LP是否在末尾
                                    name = check_LP_in_last()
                                    # print(name)

                                    # 检查一级系统并索引到不同的检测方式
                                    check_first_system_name(name)
                                    # print(name)

                                    # 通用词汇检查
                                    check_by_com_word(name)

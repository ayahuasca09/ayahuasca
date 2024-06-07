import openpyxl
import json
import sys
from os.path import abspath, dirname
import os
from pprint import pprint
import re

# 文件所在目录
py_path = ""
if hasattr(sys, 'frozen'):
    py_path = dirname(sys.executable)
elif __file__:
    py_path = dirname(abspath(__file__))

"""获取json文件路径"""
json_name = '导入规则.json'
json_path = os.path.join(py_path, json_name)

"""*****************功能检测区******************"""

"""字符串长度检查"""


def check_by_str_length(str1, length):
    if len(str1) > length:
        if "_" in str1:
            print(name + "：描述后缀名过长，限制字符数为" + str(length))
        else:
            print(name + "：通过_切分的某个单词过长，每个单词长度不应超过10个字符，需要再拆分")


"""按照json配置检查"""


def check_by_json(system_dict):
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
            ")" + "_"

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


"""检查系统名是否正确"""


def check_first_system_name(js_dict):
    flag = 0
    for key in js_dict:
        # print(key)
        if key == word_list[0]:
            flag = 1
            break
    if flag == 1:
        return js_dict[word_list[0]]
    else:
        print(name + "：一级系统名称（如Amb、Char）等有误，请检查拼写")


"""检查LP是否在末尾"""


def check_LP_in_last():
    if "LP" in name:
        # print("LP")
        # .*代表所有字符 _LP$代表以_LP结尾
        pattern = r".*_LP$"
        result = re.match(pattern, name)
        if result != None:
            pass
        else:
            print(name + "：_LP应放在最末尾")


"""测试名称"""
name = "Char_Skill_C01_Focus_Ready2_LP"
word_list = name.split("_")
is_title = True
for word in word_list:
    """判定_的每个都不能超过10个"""
    check_by_str_length(word, 10)
    """判定_的每个开头必须大写"""
    # 只需要检查第一个字符，因为istitle会导致检查字符串只能首字母大写
    result = word[0].istitle()
    if result == False:
        is_title = False
        break
if is_title == False:
    print(name + "：通过”_“分隔的每个单词开头都需要大写")

# 一级系统名称
system_name = ""

"""*****************主程序处理******************"""
# 打开json文件并转为python字典
with open(json_path, 'r') as jsfile:
    js_dict = json.load(jsfile)
    # pprint(js_dict)
    # {'Amb': {'Gen': {'description': {'length': 25}},
    #          'X00': {'child': ['MgSt', 'Crem'],
    #                  'description': {'length': 25},
    #                  'name': 'A'}},
    #  'Char': {'module': '((Skill)|(Foley)|(Mov))',
    #           'name': 'C',
    #           'property': '((Atk\\d*)|(Hook)|(SkyAtk)|(Battle)|(Dodge)|(Counter)|
    #           (Execu\\d*)|(Focus)|(Gen)|(Hit)|(Death)|(Strafe)|(Ult))'}}

    # 检查系统名是否正确
    system_dict = check_first_system_name(js_dict)
    # pprint(system_dict)
    # {'module': '((Skill)|(Foley)|(Mov))',
    #  'name': 'C',
    #  'property': '((Atk\\d*)|(Hook)|(SkyAtk)|(Battle)|(Dodge)|(Counter)|
    #  (Execu\\d*)|(Focus)|(Gen)|(Hit)|(Death)|(Strafe)|(Ult))'}

    # 检查LP是否在末尾
    check_LP_in_last()

    # check_by_json(system_dict)
    check_by_json(system_dict)

# check_LP_in_last()

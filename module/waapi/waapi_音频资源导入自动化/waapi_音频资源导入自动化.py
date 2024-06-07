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

"""测试名称"""
name = "Char_C03_Skill_Execu1_Pre_Hook"
name_list = name.split("_")

"""*****************功能检测区******************"""
"""检查系统名是否正确"""


def check_first_system_name(js_dict):
    flag = 0
    for key in js_dict:
        # print(key)
        if key == name_list[0]:
            flag = 1
    if flag == 1:
        pass
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

    check_first_system_name(js_dict)

# check_LP_in_last()

"""判定_的每个开头都要大写"""

"""判定_的每个都不能超过10个"""

"""长词需要使用缩写"""

"""正则规则列表"""
pattern_list = []

char = "Char"

"""Char检查规则"""
# 以Char开头,命名分组以供输出(?P<正则>)
pattern = ("(?P<first>"
           "^" + char +
           "_"
           # 角色名称：C01，C02……
           "C\\d{2,4}"
           ")"
           "_"
           # 所属哪个模块
           "((Skill)|(Foley))"
           "_"
           # 所属哪个Actor-Mixer,*d代表出现0或多次数字
           "((Atk\\d*)|(Hook)|(SkyAtk)|(Battle)|"
           "(Dodge)|(Counter)|"
           "(Execu\\d*)|"
           "(Focus)|"
           "(Gen)|"
           "(Hit)|(Death)|"
           "(Strafe)|"
           "(Ult))"
           "_"
           # 最后的状态描述，不能超过20个
           ".*"
           )
pattern_list.append(pattern)
# result = re.search(pattern, name)
# print(result)
# data_list = re.finditer(pattern, name)
# for item in data_list:
#     item_dict = item.groupdict()
#     pprint(item_dict)

# print(name_list)

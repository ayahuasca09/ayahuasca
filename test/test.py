import os
from os.path import abspath, dirname
import sys
import subprocess
import openpyxl
import re

"""输出测试"""
# print("aa")

"""获取当前.py所在路径测试"""
# # # 绝对路径
# # print(abspath(__file__))
# # 文件所在目录
# py_path = dirname(abspath(__file__))
# # 文件拼接
# file_name = 'test.txt'
# file_path = os.path.join(py_path, file_name)
# with open(file_path, 'r', encoding='utf-8') as file:
#     content = file.read()
# print(content)
# os.system(file_path)

"""打开文件操作测试"""
# os.system("F:/pppppy/SP/test/test.txt")

"""查找值对应的键"""
# args = {
#     "Gen_Word": {
#         "Lit": ["Light", "Small"],
#         "Mid": ["Middle", "Medium"],
#         "Hvy": ["Heavy", "Large", "Big"],
#         "Far": ["Distant"],
#         "Med": ["Medium"],
#         "Cls": ["Close", "Near"],
#         "VeryCls": ["VeryClose"],
#         "LP": ["Lp", "Loop", "lp"],
#         "Grd": ["Ground"],
#         "Death": ["Die", "Dying"],
#         "Sucs": ["Succeed", "Success"],
#         "Fail": ["False", "Lose", "Failure"]
#     }
# }
# for key in args['Gen_Word']:
#     for value in args['Gen_Word'][key]:
#         if value == "Light":
#             print(key)

"""测试获取字典的最后一位"""
# my_dict = {
#     "first_key": 'first_value',
#     "second_key": "second_value",
#     "third_key": "third_value",
# }
# print("last key : ", list(my_dict.keys())[-1])
# print("last value : ", my_dict.get(list(my_dict.keys())[-1]))


"""通过字符串调用函数"""
# def check_by_Char():
#     print("Char")
#     return True
# type = "Char"
# func_name = "check_by_" + type
# result = eval(func_name)()
# print(result)

"""在函数中改变全局变量的值"""
# a = 100
# def change_value():
#     global a
#     a = 50
# change_value()
# print(a)


"""多条件判断测试"""
# max = [3, 5, 7, 6, 8, 4]
# if 3 in max:
#     print("aaa")

"""全局变量测试"""
# def get_imax():
#     global imax
#     imax = 6
#
#
# for i in range(1, 10):
#     imax = 3
#     get_imax()
# print(imax)

"""测试exe控制命令执行"""
# command = (r"S:\Ver_1.0.0\Editor\Engine\Binaries\Win64\UnrealEditor.exe S:\Ver_1.0.0\Project\SilverPalace.uproject "
# #            r"-run=WwiseReconcileCommandlet -modes=all")
# # subprocess.run(command, capture_output=True, text=True, check=True, shell=True)

"""以_分割字符串，保存为一个列表，并去除列表中的最后一个元素，组成一个新字符串"""
# # 原始字符串
# original_string = "这是一个_测试_字符串_示例_42"
#
# # 以_分割字符串并保存为列表
# string_list = original_string.split('_')
#
# # 去除列表中的最后一个元素
# string_list.pop()
#
# # 组成一个新字符串
# new_string = '_'.join(string_list)
#
# print(new_string)

"""表格中的正则表达式提取测试"""

# # 打开 Excel 文件
# workbook = openpyxl.load_workbook(r"F:\pppppy\SP\files\工作簿1.xlsx")
#
# # 选择工作表（假设是第一个工作表）
# sheet = workbook.active
#
# # 获取第一行第一列的值
# pattern_value = sheet.cell(row=1, column=1).value
#
# # 打印模式
# print(f"Pattern from Excel: {pattern_value}")
#
# # 将双反斜杠替换为单反斜杠
# pattern_value = pattern_value.replace('\\\\', '\\')
#
# # 编译正则表达式模式
# try:
#     pattern = re.compile(pattern_value)
# except re.error as e:
#     print(f"Invalid regex pattern: {e}")
#     exit(1)
#
# # 检查字符串 "C01" 是否匹配
# match = pattern.fullmatch("C01")
#
# if match:
#     print("字符串 'C01' 匹配模式。")
# else:
#     print("字符串 'C01' 不匹配模式。")


"""event结构找父级测试"""

#
# def find_longest_key_in_dict(dict1, string):
#     # 将字符串拆分为列表
#     string_list = string.split('_')
#
#     # 过滤符合条件的键
#     valid_keys = [
#         key for key in dict1.keys()
#         if all(part in string_list for part in key.split('_'))
#     ]
#
#     if valid_keys:
#         # 找出最长的键
#         longest_key = max(valid_keys, key=len)
#         return longest_key
#     else:
#         return None
#
#
# # 示例字典和字符串
# dict1 = {
#     'key_one': 1,
#     'key_two': 2,
#     'key_three': 3
# }
#
# string = 'key_one_key_two_key_three_key_four'
#
# # 调用函数并打印结果
# longest_key = find_longest_key_in_dict(dict1, string)
# print(longest_key)

"""输入检查"""


# python输入检查
# 1.必须为数字和以,为间隔，例如1，4，2，6
# 2.通过,分隔的数字不能重复
def is_valid_input(input_string):
    # 删除空格
    input_string = input_string.replace(" ", "")

    # 检查是否仅包含数字和逗号
    if not all(c.isdigit() or c == ',' for c in input_string):
        return False

    # 拆分字符串并转换为整数列表
    try:
        numbers = list(map(int, input_string.split(',')))
    except ValueError:
        return False

    # 检查是否有重复
    if len(numbers) != len(set(numbers)):
        return False

    return True


# 示例用法
user_input = input("请输入数字，使用逗号分隔：")
if is_valid_input(user_input):
    print("输入有效。")
else:
    print("输入无效，请确保输入为数字且不重复。")

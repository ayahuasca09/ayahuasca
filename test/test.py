import os
from os.path import abspath, dirname
import sys
import subprocess

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

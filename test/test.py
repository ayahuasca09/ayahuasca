import os
from os.path import abspath, dirname

"""输出测试"""
# print("aa")

"""获取当前.py所在路径测试"""
# 绝对路径
# print(abspath(__file__))
# 文件所在目录
print(dirname(abspath(__file__)))

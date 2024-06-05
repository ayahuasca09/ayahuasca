import os
from os.path import abspath, dirname
import sys

"""输出测试"""
# print("aa")

"""获取当前.py所在路径测试"""
# # 绝对路径
# print(abspath(__file__))
# 文件所在目录
py_path = dirname(abspath(__file__))
# 文件拼接
file_name = 'test.txt'
file_path = os.path.join(py_path, file_name)
with open(file_path, 'r', encoding='utf-8') as file:
    content = file.read()
print(content)
os.system(file_path)

"""打开文件操作测试"""
# os.system("F:/pppppy/SP/test/test.txt")


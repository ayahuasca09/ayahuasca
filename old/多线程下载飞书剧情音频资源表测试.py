# xml读取库
import shutil
import os
import sys
from os.path import abspath, dirname
import openpyxl
import re
# xml写入库
from xml.dom.minidom import Document, parse
from pprint import pprint
from concurrent.futures import ThreadPoolExecutor

# 自定义库
import comlib.excel_h as excel_h
import comlib.config as config
import comlib.oi_h as oi_h
import comlib.csv_h as csv_h
import comlib.exe_h as exe_h
import comlib.cloudfeishu_h as cloudfeishu_h

"""****************数据获取******************"""

# 获取文件所在目录
py_path = ""
if hasattr(sys, 'frozen'):
    py_path = dirname(sys.executable)
elif __file__:
    py_path = dirname(abspath(__file__))

"""****************表格获取输入****************"""

"""检查输出的资源表数字索引是否正确"""


# python输入检查
# 1.必须为数字和以,为间隔，例如1，4，2，6
# 2.通过,分隔的数字不能重复
def is_valid_input(input_string, media_excel_token_keys):
    # 删除空格
    input_string = input_string.replace(" ", "")

    # 检查是否仅包含数字和逗号
    if not all((c.isdigit() or c == ',') for c in input_string):
        print("输入错误，请确保输入的只有数字且以,隔开")
        return False

    # 拆分字符串并转换为整数列表
    try:
        numbers = list(map(int, input_string.split(',')))
    except ValueError:
        print("输入错误，请确保输入的只有数字且以,隔开")
        return False

    # 检查是否有重复
    if len(numbers) != len(set(numbers)):
        print("输入错误，请输无重复的数字")
        return False

    for numbers in numbers:
        if numbers > len(media_excel_token_keys):
            print("输入错误，请输入上述表中有的数字")
            return False

    return True


def input_show(media_excel_token_dict):
    media_excel_token_keys = list(media_excel_token_dict.keys())
    print("--------------------------")
    print("输入数字对应资源表类型如下：")
    print("输入值若为all，代表更新所有资源表内容")
    print("若需获取多个资源表，以英文字符,(逗号）隔开，输入完毕后按下回车即可")
    print("输入案例：1,3,5,6")

    for i in range(len(media_excel_token_keys)):
        print("{}、".format(i + 1) + str(media_excel_token_keys[i]))
    print("--------------------------")

    temp = ''
    flag = 1
    while flag:
        temp = input("请输入要更新资源的音效表数字:")
        if temp == "all":
            for media_excel_name in media_excel_token_dict:
                print(media_excel_name + "：音频资源表更新")
                excel_token = media_excel_token_dict[media_excel_name]
                cloudfeishu_h.download_cloud_sheet(excel_token,
                                                   os.path.join(py_path, "Excel", media_excel_name) + '.xlsx')
        elif temp == "no":
            break
        elif not is_valid_input(temp, media_excel_token_keys):
            print("")
            print("请重新输入⬇")
            continue
        flag = 0
    input_digi_list = list(map(int, temp.split(',')))
    digi_meidia_dict = {index + 1: value for index, value in enumerate(media_excel_token_keys)}
    # 示例调用 multi_thread_download 函数
    multi_thread_download(input_digi_list, digi_meidia_dict, media_excel_token_dict, py_path)


# 定义一个函数来下载媒体资源表
def download_media_excel(media_excel_name, excel_token, py_path):
    # 打印当前正在更新的资源表名称
    print(f"{media_excel_name}：音频资源表更新")

    # 下载云端表格到本地指定路径
    cloudfeishu_h.download_cloud_sheet(
        excel_token,
        os.path.join(py_path, "Excel", f"{media_excel_name}.xlsx")
    )



# 主函数，用于处理输入的数字列表并下载相应的资源表
def multi_thread_download(input_digi_list, digi_meidia_dict, media_excel_token_dict, py_path):
    tasks = []  # 用于存储所有任务的列表

    # 使用 ThreadPoolExecutor 来并行执行任务
    with ThreadPoolExecutor() as executor:
        # 遍历输入的数字列表
        for i in input_digi_list:
            # 检查当前数字是否在媒体字典中
            if i in digi_meidia_dict:
                media_excel_name = digi_meidia_dict[i]  # 获取对应的媒体表名

                # 检查媒体表名是否在媒体表令牌字典中
                if media_excel_name in media_excel_token_dict:
                    excel_token = media_excel_token_dict[media_excel_name]  # 获取对应的表令牌

                    # 提交下载任务到线程池
                    # 第一个参数是函数名，后面参数是函数参数
                    task = executor.submit(download_media_excel, media_excel_name, excel_token, py_path)
                    tasks.append(task)  # 将任务添加到任务列表中

    # 等待所有任务完成（可选）
    for task in tasks:
        task.result()  # 调用 result() 确保任务完成，并捕获可能的异常


# 所有音频资源表的映射
media_excel_token_dict = config.es_sheet_token_dict
input_show(media_excel_token_dict)

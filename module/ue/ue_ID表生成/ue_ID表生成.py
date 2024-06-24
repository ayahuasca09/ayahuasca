import unreal
import openpyxl

# from waapi import WaapiClient, CannotConnectToWaapiException


# with WaapiClient() as client:
#     print("a")

"""ID表内容写入"""
# 获取ID表路径
# 打开工作簿
excel_path = r"F:\pppppy\SP\module\ue\ue_ID表生成\Audio.xlsx"
wb = openpyxl.load_workbook(excel_path)
# 获取工作表
sheet = wb['audio']

# 读表添加数据

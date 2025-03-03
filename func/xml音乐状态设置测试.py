import openpyxl
import sys
from os.path import abspath, dirname
import os
from pprint import pprint
from waapi import WaapiClient
import shutil

import comlib.cloudfeishu_h as cloudfeishu_h
import comlib.oi_h as oi_h
import comlib.config as config
import comlib.excel_h as excel_h
import comlib.waapi_h as waapi_h
import comlib.xml_h as xml_h
import comlib.exe_h as exe_h

"""根目录获取"""


def get_py_path():
    py_path = ""
    if hasattr(sys, 'frozen'):
        py_path = dirname(sys.executable)
    elif __file__:
        py_path = dirname(abspath(__file__))
    return py_path


# 打包版路径
root_path = get_py_path()

"""xml数据"""
# 存放xml表的路径
xml_path = os.path.join(root_path, 'Xml')

# 从wwise拷贝的wwu并转化为xml表
music_wwu_path = os.path.join(config.wwise_path, "Interactive Music Hierarchy")
mus_wwu_dict, _ = oi_h.get_type_file_name_and_path('.wwu', music_wwu_path)
for mus_wwu in mus_wwu_dict:
    mus_wwu_path = mus_wwu_dict[mus_wwu]
    # print(mus_wwu_path)
    mus_xml_name = mus_wwu.replace('.wwu', '.xml')
    oi_h.copy_and_rename_file(mus_wwu_path, xml_path, mus_xml_name, True)

"""*******************主程序*******************"""
xml_file_name = 'Mus_Test.xml'

# 一个xml文件的路径
xml_file_path = os.path.join(xml_path, xml_file_name)

"""读取显示一个xml表的信息"""
# xml_h.parse_xml(xml_file_path)


"""add generic path:添加一个state group"""
xml_data = """
<ArgumentList>
            <ArgumentRef Name="Process_State" ID="{3559DD54-C15C-45FE-8970-4D70CE7B75FF}"/>
          </ArgumentList>
          <EntryList>
            <Entry>
              <Path>
                <PathElementRef Name="Map" ID="{793A0FB5-4323-4D7F-AD37-3735DD0DC8A3}"/>
              </Path>
              <AudioNodeInfo>
                <AudioNodeRef Name="Mus_Map" ID="{CCADF01C-395C-4083-9647-CC6A627A69B3}" WorkUnitID="{554FEE44-2205-4672-93D2-22EBA2351C67}" Platform="Linked"/>
              </AudioNodeInfo>
            </Entry>
            <Entry>
              <Path>
                <PathElementRef Name="Story" ID="{AE24F6AC-A8E4-4A0F-8D5E-F76EEF56BAF2}"/>
              </Path>
              <AudioNodeInfo>
                <AudioNodeRef Name="Mus_Story" ID="{DBC7FED1-FAAC-467C-9D81-7C51EE4DE9EF}" WorkUnitID="{FBC01D4B-1286-41D3-AADB-2F236DF7EEDF}" Platform="Linked"/>
              </AudioNodeInfo>
            </Entry>
          </EntryList>
"""

"""在已有state group下添加一条路径"""
# xml_data = '''
# <Entry>
# 				<Path>
# 					<PathElementRef Name="Story" ID="{AE24F6AC-A8E4-4A0F-8D5E-F76EEF56BAF2}"/>
# 				</Path>
# 				<AudioNodeInfo>
# 					<AudioNodeRef Name="Mus_Story" ID="{DBC7FED1-FAAC-467C-9D81-7C51EE4DE9EF}" WorkUnitID="{FBC01D4B-1286-41D3-AADB-2F236DF7EEDF}" Platform="Linked"/>
# 				</AudioNodeInfo>
# </Entry>
# '''
# #
# # 创建新的XML元素
# new_entry = xml_h.create_element_from_xml_string(xml_data)
#
# # 将新元素插入到指定的XML文件中
# xml_h.insert_into_xml(xml_file_path, './/EntryList', new_entry)
#
# # 输出修改后的xml结构
# # xml_h.parse_xml(xml_file_path)
#
# # 关闭wwise工程
# exe_h.close_exe("Wwise.exe")
#
# # 将文件改名并写回原Wwise
# wwu_file_name = xml_file_name.replace('.xml', '.wwu')
# oi_h.copy_and_rename_file(xml_file_path, music_wwu_path, wwu_file_name, True)
#
# # 重启wwise工程
# exe_h.open_exe(exe_h.wproj_path)

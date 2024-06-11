import config
import os
from pprint import pprint
import re

"""查找目录下特定后缀的文件"""


# dir_path：查找目录 file_type：文件后缀名
def find_all_files_by_type(dir_path, file_type):
    file_names = []

    for i in os.walk(dir_path):
        file_names.append(i)

    pprint("输出文件夹下的文件名：")
    file_name_list = file_names[0][2]
    # pprint(file_name_list)

    # 提取规则：只提取某type类文件
    file_name_type_list = []
    for i in file_name_list:
        if file_type in i:
            # print(i)
            file_name_type_list.append(i)
    # pprint(file_name_type_list)


# 测试
# find_all_files_by_type(r'F:\pppppy\SP\module\excel', '.xlsx')


# 获取\Content\Audio\GeneratedSoundBanks\Windows\Event下的json文件路径
def oi_get_json_filesname():
    # 含有json的文件列表
    file_json_list = []
    file_json_dict = {}

    # 遍历文件夹下的所有子文件
    # 绝对路径，子文件夹，文件名
    for root, dirs, files in os.walk(config.bnk_path):
        file_dict = {}
        file_dict['name'] = files
        file_dict['path'] = root
        #  {'name': [
        #            'Play_Skill_Alf_EnterBattle.bnk',
        #            'Play_Skill_Alf_EnterBattle.json'],
        #   'path': 'S:\\spgame\\Project\\Content\\Audio\\GeneratedSoundBanks\\Windows\\Event\\26'},

        # 提取规则：只提取json文件
        for json_root, json_dirs, json_files in os.walk(root):
            # pprint(json_root)
            for json_file in json_files:
                if ".json" in json_file:
                    # pprint(json_file)
                    json_path = os.path.join(json_root, json_file)
                    # pprint(json_path)
                    file_json_dict['name'] = json_file
                    file_json_dict['path'] = json_path
                    file_json_list.append(file_json_dict)

    return file_json_list
    # [{'name': 'Play_VO_Munin_07.json',
    #   'path': 'S:\\spgame\\Project\\Content\\Audio\\GeneratedSoundBanks\\Windows\\Event\\English\\28\\Play_VO_Munin_07.json'}]

# 调用范例
# file_json_list = oi_get_json_filesname()
# pprint(file_json_list)

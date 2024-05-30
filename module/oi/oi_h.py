import config
import os
from pprint import pprint
import re


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
            pprint(json_root)
            for json_file in json_files:
                if ".json" in json_file:
                    # pprint(json_file)
                    json_path = os.path.join(json_root, json_file)
                    # pprint(json_path)
                    file_json_dict['name'] = json_file
                    file_json_dict['path'] = json_path
                    file_json_list.append(file_json_dict)

    return file_json_list


# 调用范例
# file_json_list = oi_get_json_filesname()
# pprint(file_json_list)

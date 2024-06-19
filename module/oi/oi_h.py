import config
import os
from pprint import pprint
import re
import shutil

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

"""遍历文件目录获取文件名称和路径
"""


def get_type_file_name_and_path(file_type, dir_path):
    file_dict = {}
    # 遍历文件夹下的所有子文件
    # 绝对路径，子文件夹，文件名
    for root, dirs, files in os.walk(dir_path):
        # {'name': ['1111.akd', 'Creature Growls 4.akd', 'Sonic Salute_005_2_20.akd'],
        #  'path': 'S:\\chen.gong_DCC_Audio\\Audio\\SilverPalace_WwiseProject\\Originals\\SFX'}
        for file in files:
            if file_type in file:
                file_dict[file] = os.path.join(
                    root, file)

    return file_dict


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

# 复制文件到指定文件夹并重命名
# uu = shutil.copy2(r"F:\pppppy\SP\module\waapi\waapi_音频资源导入自动化\Media_Temp.wav"
#                   , os.path.join(r"F:\pppppy\SP\module\waapi\waapi_音频资源导入自动化\New_Media", 'Media_New' + ".wav"))
# print(uu)

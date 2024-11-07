import module.config as config
import os
from pprint import pprint
import sys
from os.path import abspath, dirname
import re
import module.excel.excel_h as excel_h
import module.waapi.waapi_h as waapi_h
import module.json.json_h as json_h

is_pass = True

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

"""遍历文件目录获取文件名称和路径"""


def get_type_file_name_and_path(file_type, dir_path):
    file_dict = {}
    file_list = []
    # 遍历文件夹下的所有子文件
    # 绝对路径，子文件夹，文件名
    for root, dirs, files in os.walk(dir_path):
        # {'name': ['1111.akd', 'Creature Growls 4.akd', 'Sonic Salute_005_2_20.akd'],
        #  'path': 'S:\\chen.gong_DCC_Audio\\Audio\\SilverPalace_WwiseProject\\Originals\\SFX'}
        for file in files:
            if file_type in file:
                new_dict = {}
                file_dict[file] = os.path.join(
                    root, file)
                new_dict[file] = os.path.join(
                    root, file)
                file_list.append(new_dict)

    return file_dict, file_list


r"""获取Content\Audio\GeneratedSoundBanks\Windows\Event下的json文件路径"""


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


"""获取文件所在目录"""


def get_py_path():
    py_path = ""
    if hasattr(sys, 'frozen'):
        py_path = dirname(sys.executable)
    elif __file__:
        py_path = dirname(abspath(__file__))
    return py_path


"""log捕获"""


def print_warning(warn_info):
    print("[warning]" + warn_info)


"""报错捕获"""


def print_error(error_info):
    global is_pass
    is_pass = False
    print("[error]" + error_info)
    return False


"""检测字符串是否含有中文"""


def check_is_chinese(string):
    for ch in string:
        if u'\u4e00' <= ch <= u'\u9fff':
            return True
    return False


"""字符串长度检查"""


def check_by_str_length(str1, length, name):
    if len(str1) > length:
        if "_" in str1:
            print_error(name + "：描述后缀名过长，限制字符数为" + str(length))
        else:
            print_error(name + "：通过_切分的某个单词过长，每个单词长度不应超过10个字符，需要再拆分")


"""分隔符的大小写和长度检查"""


def check_by_length_and_word(name):
    # pprint(cell_sound.value)
    word_list = name.split("_")
    is_title = True
    for word in word_list:
        """判定_的每个都不能超过10个"""
        check_by_str_length(word, 10, name)
        """判定_的每个开头必须大写"""
        if len(word) > 0:
            # 只需要检查第一个字符，因为istitle会导致检查字符串只能首字母大写
            if word[0].istitle() == False and word[0].isdigit() == False:
                is_title = False
                break
    if is_title == False:
        print_error(name + "：通过”_“分隔的每个单词开头都需要大写")
    if word_list:
        return word_list[0]
    else:
        return None


"""检查LP是否在末尾"""


def check_LP_in_last(name):
    if "LP" in name:
        # print("LP")
        # .*代表所有字符 _LP$代表以_LP结尾
        pattern = r".*_LP$"
        result = re.match(pattern, name)
        if result:
            new_name = re.sub(r"_LP$", "", name)
            return new_name
        else:
            print_error(name + "：_LP应放在最末尾")
    else:
        return name


"""音频命名规范基础检查"""


def check_name_all(name):
    global is_pass
    is_pass = True
    first_name = check_by_length_and_word(name)
    no_LP_name = check_LP_in_last(name)
    return first_name, no_LP_name, is_pass


"""检查名称是否为随机"""


def check_is_random(name):
    is_random = re.search(r"(_R\d{2,4})$", name)
    if is_random:
        return True
    else:
        return False


# 调用范例
# file_json_list = oi_get_json_filesname()
# pprint(file_json_list)

"""复制文件到指定文件夹并重命名"""
# uu = shutil.copy2(r"F:\pppppy\SP\module\waapi\waapi_音频资源导入自动化\Media_Temp.wav"
#                   , os.path.join(r"F:\pppppy\SP\module\waapi\waapi_音频资源导入自动化\New_Media", 'Media_New' + ".wav"))
# print(uu)

"""修改文件名称"""
# os.rename(r"C:\Users\happyelements\Desktop\面试问题补充.txt",r"C:\Users\happyelements\Desktop\面试.txt")

import openpyxl
import json
import sys
from os.path import abspath, dirname
import os
from pprint import pprint
import re
from waapi import WaapiClient
import shutil
from openpyxl.cell import MergedCell

# 文件所在目录
py_path = ""
if hasattr(sys, 'frozen'):
    py_path = dirname(sys.executable)
elif __file__:
    py_path = dirname(abspath(__file__))

"""获取.xlsx文件"""

file_names = []
for i in os.walk(py_path):
    file_names.append(i)
# pprint("输出文件夹下的文件名：")
file_name_list = file_names[0][2]

"""状态类通用后缀列表"""
list_state_suffix = ['State', 'Type']

"""事件描述查重列表"""
event_desc_list = []

"""*****************功能检测区******************"""
"""检查是否为合并单元格"""


def check_is_mergecell(cell):
    if isinstance(cell, MergedCell):  # 判断该单元格是否为合并单元格
        for merged_range in sheet.merged_cells.ranges:  # 循环查找该单元格所属的合并区域
            if cell.coordinate in merged_range:
                # 获取合并区域左上角的单元格作为该单元格的值返回
                cell = sheet.cell(row=merged_range.min_row, column=merged_range.min_col)
                break
    return cell.value


"""检查是否使用通用词汇"""


def check_by_com_suffix(suffix_name):
    if suffix_name in list_state_suffix:
        return
    else:
        print_error(cell_sound.value + "后缀错误或未添加至通用后缀列表")


"""报错捕获"""


def print_error(error_info):
    global is_pass
    is_pass = False
    print("[error]" + error_info)


"""检测字符串是否含有中文"""


def check_is_chinese(string):
    for ch in string:
        if u'\u4e00' <= ch <= u'\u9fff':
            return True
    return False


"""字符串长度检查"""


def check_by_str_length(str1, length):
    if len(str1) > length:
        if "_" in str1:
            print_error(cell_sound.value + "：描述后缀名过长，限制字符数为" + str(length))
        else:
            print_error(cell_sound.value + "：通过_切分的某个单词过长，每个单词长度不应超过10个字符，需要再拆分")


"""获取表格中的事件描述和状态所在的列"""


def get_descrip_and_status_column():
    require_module_column = None
    second_module_column = None
    require_name_column = None
    status_column = None
    if list(sheet.rows)[0]:
        for cell in list(sheet.rows)[0]:
            if cell.value:
                if 'Group名称' == str(cell.value):
                    require_module_column = cell.column
                    # print(require_module_column)
                elif 'Group描述' == str(cell.value):
                    second_module_column = cell.column
                    # print(second_module_column)
                elif '值' == str(cell.value):
                    require_name_column = cell.column
                    # print(require_name_column)
                elif '值描述' == str(cell.value):
                    status_column = cell.column
                    # print(status_column)

    return require_module_column, second_module_column, require_name_column, status_column


"""分隔符的大小写和长度检查"""


def check_by_length_and_word(name):
    # pprint(cell_sound.value)
    word_list = name.split("_")
    is_title = True
    for word in word_list:
        """判定_的每个都不能超过10个"""
        check_by_str_length(word, 10)
        """判定_的每个开头必须大写"""
        if len(word) > 0:
            # 只需要检查第一个字符，因为istitle会导致检查字符串只能首字母大写
            if word[0].istitle() == False and word[0].isdigit() == False:
                is_title = False
                break
    if is_title == False:
        print_error(name + "：通过”_“分隔的每个单词开头都需要大写")
    if word_list:
        return word_list[-1]
    else:
        return None


with WaapiClient() as client:
    """*****************Wwise功能区******************"""
    """创建新的随机容器"""


    def create_state(state_name):
        # 查找该rnd是否已存在
        flag = 0
        # 去除随机容器数字
        rnd_name = re.sub(r"(_R\d{2,4})$", "", media_name)

        rnd_path = ""
        rnd_container_list, rnd_id, _ = find_obj(
            {'waql': ' "%s" select descendants where type = "RandomSequenceContainer" ' % os.path.join(
                wwise_dict['Root'], system_name)})
        # 找到重名的
        for rnd_container_dict in rnd_container_list:
            if rnd_container_dict['name'] == rnd_name:
                flag = 1
                rnd_path = rnd_container_dict['path']
                set_obj_notes(rnd_container_dict['id'], event_descrip)
                break
            # 没有重名但描述一致，判定为改名
            elif rnd_container_dict['notes'] == event_descrip:
                # pprint(rnd_name + "：RandomContainer已存在，将不再导入")
                flag = 2
                rnd_path = rnd_container_dict['path'].replace(rnd_container_dict['name'], rnd_name)
                change_name_by_wwise_content(rnd_container_dict['id'], rnd_name, rnd_container_dict['name'],
                                             "RandContainer")
                # Sound也要改名
                sound_container_list, sound_container_id, _ = find_obj(
                    {'waql': ' "%s" select children  ' % rnd_container_dict['id']})
                for sound_cotainer_dict in sound_container_list:
                    wav_tail = re.search(r"(_R\d{2,4})$", sound_cotainer_dict['name'])
                    sound_name = rnd_name
                    if wav_tail:
                        sound_name = rnd_name + wav_tail.group()
                    change_name_by_wwise_content(sound_cotainer_dict['id'], sound_name, sound_cotainer_dict['name'],
                                                 "Sound")
                    # Media路径更改
                    new_media_path = sound_cotainer_dict['originalWavFilePath'].replace(sound_cotainer_dict['name'],
                                                                                        sound_name)
                    # Media重命名
                    os.rename(sound_cotainer_dict['originalWavFilePath'],
                              new_media_path)
                    # 在容器中重新导入media
                    import_media_in_rnd(new_media_path, sound_name, rnd_path)
                    print_warning(sound_cotainer_dict['originalWavFilePath'] + "(Media)改名为：" + new_media_path)

                    # 语音的话需要删除容器内的引用
                    if system_name == "VO":
                        sound_language_list, _, _ = find_obj(
                            {'waql': ' "%s" select children  ' %
                                     sound_cotainer_dict['id']})
                        if sound_language_list:
                            for sound_language_dict in sound_language_list:
                                if sound_cotainer_dict['name'] in sound_language_dict['name']:
                                    args = {
                                        "object": "%s" % sound_language_dict['id']
                                    }
                                    client.call("ak.wwise.core.object.delete", args)
                                    print_warning(
                                        sound_language_dict['name'] + "：容器引用删除")

                break

        # 不存在则创建
        if flag == 0:
            # 查找所在路径
            rnd_parent_id = find_obj_parent(rnd_name,
                                            {
                                                'waql': ' "%s" select descendants,this where type = "ActorMixer" ' %
                                                        os.path.join(
                                                            wwise_dict['Root'], system_name)})
            if rnd_parent_id:

                # 创建的rnd属性
                args = {
                    # 选择父级
                    "parent": rnd_parent_id,
                    # 创建类型机名称
                    "type": "RandomSequenceContainer",
                    "name": rnd_name,
                    "notes": event_descrip,
                    "@RandomOrSequence": 1,
                    "@NormalOrShuffle": 1,
                    "@RandomAvoidRepeatingCount": 1
                }
                rnd_container_object = client.call("ak.wwise.core.object.create", args)
                # 查找新创建的容器的路径
                _, _, rnd_path = find_obj(
                    {'waql': ' "%s"  ' % rnd_container_object['id']})
                if rnd_path:
                    print_warning(rnd_name + "：RandContainer创建")
                else:
                    print_error(rnd_name + "：RandContainer未能成功创建")

                # 声音容器创建及导入
                import_sound_sfx_and_media(media_name, rnd_path)

        return rnd_path, rnd_name


    """媒体资源导入的总流程"""


    def create_wwise_content(state_name, system_name):

        # 随机容器创建
        rnd_path, rnd_name = create_state(state_name)

        # 事件自动生成
        create_event(rnd_name, rnd_path)


    """*****************主程序处理******************"""
    # 撤销开始
    client.call("ak.wwise.core.undo.beginGroup")

    # 记录所有资源名称
    sound_name_list = []

    # 提取规则：只提取xlsx文件
    for i in file_name_list:
        if ".xlsx" in i:
            # 拼接xlsx的路径
            file_path_xlsx = os.path.join(py_path, i)
            # 获取xlsx的workbook
            wb = openpyxl.load_workbook(file_path_xlsx)
            # 获取xlsx的所有sheet
            sheet_names = wb.sheetnames
            # 加载所有工作表
            for sheet_name in sheet_names:
                sheet = wb[sheet_name]
                group_name, group_desc, value, value_desc = get_descrip_and_status_column()
                # 获取工作表第一行数据
                for cell in list(sheet.rows)[0]:
                    if '值' == str(cell.value):
                        # 获取音效名下的内容
                        for cell_sound in list(sheet.columns)[cell.column - 1]:
                            # 空格和中文不检测
                            if cell_sound.value:
                                if not check_is_chinese(cell_sound.value):

                                    """❤❤❤❤数据获取❤❤❤❤"""
                                    """检测表格内容"""
                                    is_pass = True

                                    """检查表格中是否有内容重复项"""
                                    if cell_sound.value in sound_name_list:
                                        print_error(cell_sound.value + "：表格中有重复项Group名称 ，请检查")
                                    else:
                                        sound_name_list.append(cell_sound.value)

                                        """❤❤❤state内容检查❤❤❤"""
                                        """state名称"""
                                        value = cell_sound.value
                                        suffix_name = check_by_length_and_word(value)

                                        """❤❤❤state group内容检查❤❤❤"""
                                        # 检测是否为合并单元格，是则读取合并单元格的内容
                                        """state group名称"""
                                        group_value = check_is_mergecell(
                                            sheet.cell(row=cell_sound.row, column=group_name))
                                        if group_value:
                                            # pprint(group_value)
                                            # 分隔符的大小写和长度检查，取后缀名
                                            suffix_name = check_by_length_and_word(group_value)
                                            # print(suffix_name)

                                            # 通用后缀检查
                                            check_by_com_suffix(suffix_name)

                                            """❤❤❤描述项检查❤❤❤"""
                                            """state group描述"""
                                            group_desc_value = check_is_mergecell(
                                                sheet.cell(row=cell_sound.row, column=group_desc))
                                            if group_desc_value:
                                                value_desc_value = sheet.cell(row=cell_sound.row,
                                                                              column=value_desc).value
                                                if value_desc_value:
                                                    # print(value_desc_value)
                                                    # 拼接Group和值的描述
                                                    """state描述"""
                                                    state_desc_value = group_desc_value + "_" + value_desc_value
                                                    # print(state_desc_value)
                                                    # state描述查重
                                                    if state_desc_value not in event_desc_list:
                                                        event_desc_list.append(state_desc_value)
                                                    else:
                                                        print_error(state_desc_value + "：表格中有重复项描述，请检查")

                                            else:
                                                print_error(value + ":没有状态描述，请补充！")
                                        else:
                                            print_error(group_value + ":没有状态描述，请补充！")

                                    # # 生成Wwise内容
                                    if is_pass:
                                        # print(state_desc_value)
                                        # create_wwise_content(cell_sound.value, system_name)

# 撤销结束
client.call("ak.wwise.core.undo.endGroup", displayName="rnd创建撤销")

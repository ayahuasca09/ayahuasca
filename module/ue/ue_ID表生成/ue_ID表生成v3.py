from waapi import WaapiClient
import os
import sys
import re
import pandas as pd
import tkinter as tk
from tkinter import messagebox
from os.path import abspath, dirname
from pprint import pprint

# 自定义库
import module.excel.excel_h as excel_h
import module.config as config
import module.oi.oi_h as oi_h
import module.waapi.waapi_h as waapi_h

# 文件所在目录
py_path = ""
if hasattr(sys, 'frozen'):
    py_path = dirname(sys.executable)
elif __file__:
    py_path = dirname(abspath(__file__))

"""excel表路径"""
excel_path = os.path.join(py_path, config.excel_dt_audio_path)
# 获取excel ID表路径
sheet, wb = excel_h.excel_get_sheet(excel_path, config.dt_audio_sheet_name)

"""csv表路径"""
csv_path = os.path.join(py_path, config.csv_dt_audio_path)

"""UE工程路径"""
ue_audio_path = config.ue_event_path

"""Wwise路径"""
event_root = config.wwise_event_path

range_min = 0
range_max = 0
# ID
temp = 0

"""ue路径转换"""


def convert_ue_event_path(old_path):
    new_path = ""
    # 获取从\Audio开始的内容
    result = re.search('(?<=\\\\Content).*', old_path).group()
    # \Audio\WwiseAudio\Events\v1\VO\VO_Game\VO_Game_C01\VO_Game_Battle_C01\AKE_Play_VO_Game_Battle_C01_01.uasset
    # /替换\
    result = result.replace("\\", "/")
    # /Audio/WwiseAudio/Events/v1/VO/VO_Game/VO_Game_C01/VO_Game_Battle_C01/AKE_Play_VO_Game_Battle_C01_01.uasset
    # 去除.uasset
    result = result.replace(".uasset", "")
    # /Audio/WwiseAudio/Events/v1/VO/VO_Game/VO_Game_C01/VO_Game_Battle_C01/AKE_Play_VO_Game_Battle_C01_01
    # 获取后缀
    result_list = result.split("/")
    # 拼接字符串
    if result_list[-1]:
        new_path = '/Script/AkAudio.AkAudioEvent\'/Game' + result + "." + result_list[-1] + '\''
        # print(result)
    return new_path


"""获取ID区间"""


def get_module_range(event_name):
    global range_min, range_max
    flag = 0
    for module_name in config.event_id_config:
        if str(module_name) in str(event_name):
            range_min = config.event_id_config[module_name]['min']
            range_max = config.event_id_config[module_name]['max']
            flag = 1
    if flag == 0:
        range_min = 0
        range_max = 0
        oi_h.print_error("找不到" + event_name + "的系统名")


# 设置事件路径
def set_event_path(insert_row):
    flag = 0
    # Event路径赋值
    for asset_path in ue_audio_path_list:
        if event_dict['name'] in asset_path:
            sheet.cell(row=insert_row, column=config.audioevent_index).value = asset_path
            flag = 1
            break
    if flag == 0:
        sheet.cell(row=insert_row, column=config.audioevent_index).value = ""


# ID生成
def create_id():
    global temp
    # 获取当前的最小序列
    # 记录累加的数
    i = range_min
    # 获取第0列
    for cell in list(sheet.columns)[0]:
        if cell.value and (type(cell.value) is int):
            # print(cell.value)
            if range_min < cell.value < range_max:
                if cell.value > i:
                    i = cell.value
    temp = i + 1


# 获取是否循环并赋值
def get_is_loop():
    is_loop = False
    # 查找事件引用的对象
    obj_list = client.call("ak.wwise.core.object.get",
                           waapi_h.waql_by_type("Action", event_dict['id']),
                           options=config.options)['return']
    obj_sub_list, _, _ = waapi_h.find_obj(obj_list)

    if obj_sub_list:
        for obj_dict in obj_sub_list:
            if is_loop:
                break
            else:
                # 为play的事件才继续查
                if obj_dict['ActionType'] == 1:
                    obj_list = client.call("ak.wwise.core.object.get",
                                           waapi_h.waql_by_type("Sound", obj_dict['Target']['id']),
                                           options=config.options)['return']
                    sound_list, _, _ = waapi_h.find_obj(obj_list)
                    # pprint(sound_list)
                    if sound_list:
                        for sound_dict in sound_list:
                            if sound_dict['IsLoopingEnabled']:
                                is_loop = True
                                break

                    # 音乐循环查询
                    obj_list = client.call("ak.wwise.core.object.get",
                                           waapi_h.waql_by_type("MusicPlaylistContainer", obj_dict['Target']['id']),
                                           options=config.options)['return']
                    sound_list, _, _ = waapi_h.find_obj(obj_list)
                    # pprint(sound_list)
                    if sound_list:
                        for sound_dict in sound_list:
                            if sound_dict['id'] in loop_musicplaylist_container_id_list:
                                is_loop = True
                                # print(event_dict['name'])
                                break

    return is_loop


# 表的值设置
def set_value():
    global temp
    create_id()
    flag = 0
    for cell in list(sheet.columns)[2]:
        if event_dict['name'] == cell.value:
            flag = 1
            # 事件描述赋值
            sheet.cell(row=cell.row, column=config.desc_index).value = event_dict['notes']
            set_event_path(cell.row)
            # 获取是否循环并赋值
            sheet.cell(row=cell.row, column=config.isloop_index).value = get_is_loop()
            break
        elif (event_dict['notes'] == sheet.cell(cell.row, column=config.desc_index).value) and event_dict['notes'] \
                and sheet.cell(cell.row, column=config.desc_index).value:
            flag = 2
            # 事件名称复制
            sheet.cell(row=cell.row, column=config.audioname_index).value = event_dict['name']
            set_event_path(cell.row)
            # 获取是否循环并赋值
            sheet.cell(row=cell.row, column=config.isloop_index).value = get_is_loop()
        # 是否循环赋值
    if flag == 0:
        # 插入为空的行
        insert_row = sheet.max_row + 1
        # pprint(insert_row)
        # rowname赋值
        sheet.cell(row=insert_row, column=config.rowname_index).value = temp
        # id编号赋值
        sheet.cell(row=insert_row, column=config.audioid_index).value = temp
        # 事件名称赋值
        sheet.cell(row=insert_row, column=config.audioname_index).value = event_dict['name']
        # 事件描述赋值
        sheet.cell(row=insert_row, column=config.desc_index).value = event_dict['notes']
        set_event_path(insert_row)
        # 获取是否循环并赋值
        sheet.cell(row=insert_row, column=config.isloop_index).value = get_is_loop()


"""复制到dcc的Audio表下"""


def copy_to_dcc_audio():
    """dcc audio表默认值填充"""
    # 遍历每一行（跳过标题行）
    for row in sheet_dcc.iter_rows(min_row=2):  # 假设第一行是标题行
        # 获取第7列、第8列和第9列的单元格
        cell_col7 = row[config.FadeDuration_index - 1]  # 第7列
        cell_col8 = row[config.FadeCurveNum_index - 1]  # 第8列
        cell_col9 = row[config.ObjectType_index - 1]  # 第9列

        # 填充第7列的空值为1000
        if cell_col7.value is None:
            cell_col7.value = 1000

        # 填充第8列的空值为0
        if cell_col8.value is None:
            cell_col8.value = 0

        # 填充第9列的空值为0
        if cell_col9.value is None:
            cell_col9.value = 0

    """先从dcc复制手动配置项到audio表"""
    # 创建一个字典来存储目标表（表2）第一列的值及其对应行的第7至9列数据
    excel2_data = {}

    # 遍历dcc的第一列，存储第7至9列的数据
    for row in sheet_dcc.iter_rows(min_row=2, max_row=sheet_dcc.max_row, max_col=config.ObjectType_index,
                                   values_only=True):
        key = row[0]  # 第一列的值作为键
        values = row[6:9]  # 获取第7到第9列的数据
        excel2_data[key] = values
    # pprint(excel2_data)

    # 遍历audio的第一列，若找到相等的值，则复制数据
    # 遍历 Excel1 的第一列，复制数据
    for row in sheet.iter_rows(min_row=2, max_col=1):
        key = row[0].value
        # print(key)
        if key in excel2_data:
            # 如果在 Excel2 中找到了匹配的键，复制数据
            for i, value in enumerate(excel2_data[key], start=config.FadeDuration_index):
                sheet.cell(row=row[0].row, column=i, value=value)
    # 保存原工作簿
    wb.save(excel_path)

    """再从audio表获取数据复制到dcc"""

    # 清空目标工作表的内容
    for row in sheet_dcc.iter_rows(min_row=2, max_row=sheet_dcc.max_row):
        for cell in row:
            cell.value = None

    # 复制所有到dcc
    for row_idx, row in enumerate(
            sheet.iter_rows(min_col=1, max_col=config.ObjectType_index, values_only=True), start=1):
        for col_idx, value in enumerate(row, start=1):
            sheet_dcc.cell(row=row_idx, column=col_idx, value=value)

    # 保存目标工作簿
    wb_dcc.save(config.excel_dcc_dt_audio_path)


"""从Wwise读Event_Info"""
with WaapiClient() as client:
    """判定音乐是否循环"""


    def find_music_lp_list():
        # PS：MusicPlaylistItem的根节点只有owner，
        # MusicPlaylistItem的其他子节点只有parent，parent为根节点
        obj_list = client.call("ak.wwise.core.object.get",
                               waapi_h.waql_from_type('MusicPlaylistItem'),
                               options=config.options)['return']
        musicplaylistitem_list, _, _ = waapi_h.find_obj(obj_list)
        # pprint(refer_list)

        # 使用一个列表保存为loop的musicplaylist container id
        loop_musicplaylist_container_id_list = []

        # 查找为loop的、非根节点的子节点
        for musicplaylistitem_dict in musicplaylistitem_list:
            if musicplaylistitem_dict['LoopCount'] == 0:
                # 获取非根子节点
                if (musicplaylistitem_dict['PlaylistItemType'] == 1):
                    # pprint(musicplaylistitem_dict)
                    # print()
                    # 获取子节点的parent，即根节点
                    root_id = musicplaylistitem_dict['parent']['id']
                    # print(root_id)
                    # 获取根节点的owner，即musicplaylist容器的id
                    obj_list = client.call("ak.wwise.core.object.get",
                                           waapi_h.waql_from_id(root_id),
                                           options=config.options)['return']
                    root_list, _, _ = waapi_h.find_obj(obj_list)

                # 获取根节点
                else:
                    # 获取根节点的owner，即musicplaylist容器的id
                    obj_list = client.call("ak.wwise.core.object.get",
                                           waapi_h.waql_from_id(musicplaylistitem_dict['id']),
                                           options=config.options)['return']
                    root_list, _, _ = waapi_h.find_obj(obj_list)
                    # pprint(root_list)
                for root_dict in root_list:
                    # 添加非重复元素，set为去重，然后再转为list
                    # 获取的是为loop的非根子节点的父级
                    loop_musicplaylist_container_id_list = list(
                        set(loop_musicplaylist_container_id_list + [root_dict['owner']['id']]))
        return loop_musicplaylist_container_id_list


    """*************主程序*************"""

    # 获取为循环的音乐列表
    loop_musicplaylist_container_id_list = find_music_lp_list()
    # pprint(loop_musicplaylist_container_id_list)

    # 存放UE的Event路径
    ue_audio_path_list = []
    # 获取ue工程下的audio并转换路径
    ue_audio_dict, _ = oi_h.get_type_file_name_and_path('.uasset', ue_audio_path)
    for ue_audio in ue_audio_dict:
        if convert_ue_event_path(ue_audio_dict[ue_audio]):
            ue_audio_path_list.append(convert_ue_event_path(ue_audio_dict[ue_audio]))
    obj_list = client.call("ak.wwise.core.object.get",
                           waapi_h.waql_by_type("Event", event_root),
                           options=config.options)['return']
    event_list, event_id, _ = waapi_h.find_obj(obj_list)
    for event_dict in event_list:
        # 事件需要勾选才生成info信息
        if event_dict['Inclusion']:
            # 获取值域范围
            get_module_range(event_dict['name'])
            if (range_min != 0) and (range_max != 0):
                set_value()

    # 同步删除Wwise中没有的事件
    # 创建Event Name列表
    extract_key = lambda d: d['name']
    event_name_list = list(map(extract_key, event_list))
    # pprint(event_name_list)
    # 获取AudioName列的内容
    for col in sheet.iter_cols(1, sheet.max_column):
        if col[0].value == 'AudioName':
            # 返回该列的值（除了标题行）
            for cell in col[1:]:
                if cell.value not in event_name_list:
                    sheet.delete_rows(cell.row, 1)
                    oi_h.print_warning(cell.value + "：在Wwise Event中未被找到，该ID及数据已删除")

wb.save(excel_path)

# excel转csv
df = pd.read_excel(excel_path)
# 指定CSV文件名和编码格式
encoding = 'utf-8'
df.to_csv(csv_path, encoding=encoding, index=False)

# 将表内容同步到dcc下的Audio表中
sheet_dcc, wb_dcc = excel_h.excel_get_sheet(config.excel_dcc_dt_audio_path, config.dt_audio_sheet_name)
copy_to_dcc_audio()

# 应用程序弹窗
root = tk.Tk()
root.withdraw()
messagebox.showinfo("Info", "ID表已更新，记得提交DCC下的Audio表")

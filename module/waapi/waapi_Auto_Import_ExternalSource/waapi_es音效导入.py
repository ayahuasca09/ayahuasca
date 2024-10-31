# xml读取库
import shutil
import pandas as pd
import os
from waapi import WaapiClient
from pprint import pprint
import sys
from os.path import abspath, dirname

# xml写入库
from xml.dom.minidom import Document, parse

# 自定义库
import module.excel.excel_h as excel_h
import module.config as config
import module.oi.oi_h as oi_h
import module.waapi.waapi_h as waapi_h
import module.cloudfeishu.cloudfeishu_h as cloudfeishu_h

"""****************数据获取******************"""
# 获取相应的excel表
sheet_mediainfo, wb_mediainfo = excel_h.excel_get_sheet(config.excel_mediainfo_path, 'Sheet1')
sheet_wwisecookie, wb_wwisecookie = excel_h.excel_get_sheet(config.excel_wwisecookie_path, 'Sheet1')

# 获取文件所在目录
py_path = ""
if hasattr(sys, 'frozen'):
    py_path = dirname(sys.executable)
elif __file__:
    py_path = dirname(abspath(__file__))

# 获取.csv文件
media_info_path = os.path.join(py_path, config.csv_mediainfo_path)
external_cookie_path = os.path.join(py_path, config.csv_wwisecookie_path)

# 获取媒体资源文件列表
wav_path = os.path.join(py_path, "New_Media")

# 获取各表token
excel_es_vo_token = config.excel_es_vo_token
excel_es_sfx_token = config.excel_es_sfx_token

# excel表标题列
excel_es_title_list = config.excel_es_title_list

"""**************写xml**************"""
# 创建一个对象
doc = Document()
# 创建根节点
root = doc.createElement("ExternalSourcesList")
root.setAttribute("SchemaVersion", "1")

# 如果保存在Wwise工程下，此Root为Wwise工程的相对路径，则不需要输入绝对路径，仅需输入.wav的名字
# 否则取文件的绝对路径
# root.setAttribute("Root", "UngeneratedExternalSources")
# 添加载Document对象上
doc.appendChild(root)

"""*****************功能检测区******************"""

"""xml创建元素"""


def xml_create_element(media_path):
    source = doc.createElement("Source")
    source.setAttribute("Path", media_path)
    source.setAttribute("Conversion", "VO")
    root.appendChild(source)


"""写入media_info excel表"""


def write_media_info_excel(vo_id, cell_sound):
    flag = 0
    for cell in list(sheet_mediainfo.columns)[0]:
        # 在media_info表中找到该ID，则替换
        if sheet_mediainfo.cell(row=cell.row, column=1).value == vo_id:
            flag = 1
            # 找到则替换，id不会改变，会改变其他内容,目前只有name可能会改
            for index, value in enumerate(list(sheet_mediainfo.rows)[0]):
                if value == "MediaName":
                    sheet_mediainfo.cell(row=cell.row, column=index + 1).value = cell_sound.value + ".wem"
                break
    # 在media_info表中未找到该ID，则新增
    if flag == 0:
        # 插入为空的行
        insert_row = sheet_mediainfo.max_row + 1
        value_dict = {
            "Name": vo_id,
            'ExternalSourceMediaInfoId': vo_id,
            'MediaName': cell_sound.value + ".wem",
            'CodecID': 4,
            'bIsStreamed': 'TRUE',
            'bUseDeviceMemory': "FALSE",
            'MemoryAlignment': 0,
            'PrefetchSize': 0
        }
        for cell in list(sheet_mediainfo.rows)[0]:
            if cell.value in value_dict:
                sheet_mediainfo.cell(row=insert_row, column=cell.column).value = value_dict[cell.value]


"""写入wwise_cookie excel表"""


def write_wwise_cookie_excel(vo_id, cell_sound, external_sound_dict):
    flag = 0
    for cell in list(sheet_wwisecookie.columns)[0]:
        # 在media_info表中找到该ID，则替换
        if sheet_wwisecookie.cell(row=cell.row, column=1).value == vo_id:
            flag = 1
            # 找到则替换，id不会改变，会改变其他内容,目前只有name可能会改
            for index, value in enumerate(list(sheet_wwisecookie.rows)[0]):
                if value == "MediaName":
                    sheet_wwisecookie.cell(row=cell.row, column=index + 1).value = cell_sound.value + ".wem"
                elif value == "ExternalSourceCookie":
                    sheet_wwisecookie.cell(row=cell.row, column=index + 1).value = external_sound_dict['shortId']
                elif value == "ExternalSourceName":
                    sheet_wwisecookie.cell(row=cell.row, column=index + 1).value = external_sound_dict['name']
                break
    # 在media_info表中未找到该ID，则新增
    if flag == 0:
        # 插入为空的行
        insert_row = sheet_wwisecookie.max_row + 1
        value_dict = {
            "Name": vo_id,
            'ExternalSourceCookie': external_sound_dict['shortId'],
            'ExternalSourceName': external_sound_dict['name'],
            'MediaInfoId': vo_id,
            'MediaName': cell_sound.value + ".wem"
        }
        for cell in list(sheet_wwisecookie.rows)[0]:
            if cell.value in value_dict:
                sheet_wwisecookie.cell(row=insert_row, column=cell.column).value = value_dict[cell.value]


"""删除相应的.wem文件"""


def delete_cancel_wem(media_name):
    _, wem_list = oi_h.get_type_file_name_and_path('.wem', config.external_output_path)
    # [{'VO_C01_15_World.wem': 'S:\\Ver_1.0.0\\Project\\Content\\Audio\\GeneratedExternalSources\\Android\\VO_C01_15_World.wem'},
    #  {'VO_C01_33_Battle.wem': 'S:\\Ver_1.0.0\\Project\\Content\\Audio\\GeneratedExternalSources\\Android\\VO_C01_33_Battle.wem'},
    #  {'VO_C02_06_Menu.wem': 'S:\\Ver_1.0.0\\Project\\Content\\Audio\\GeneratedExternalSources\\Android\\VO_C02_06_Menu.wem'},
    #  {'VO_C02_33_Battle.wem': 'S:\\Ver_1.0.0\\Project\\Content\\Audio\\GeneratedExternalSources\\Android\\VO_C02_33_Battle.wem'},]
    for wem_dict in wem_list:
        for wem_name in wem_dict:
            # print(wem_name)
            if media_name in wem_name:
                if os.path.isfile(wem_dict[wem_name]):
                    os.remove(wem_dict[wem_name])
                    oi_h.print_warning("wem文件删除：" + str(wem_dict[wem_name]))


"""移除标记为cancel的所有内容"""


def delete_cancel_content(vo_id):
    pass
    # 删除excel表的相应内容
    for cell in list(sheet_mediainfo.columns)[0]:
        # 在media_info表中找到该ID，则删除
        if sheet_mediainfo.cell(row=cell.row, column=1).value == vo_id:
            sheet_mediainfo.delete_rows(cell.row)
        if sheet_wwisecookie.cell(row=cell.row, column=1).value == vo_id:
            sheet_wwisecookie.delete_rows(cell.row)


"""自动化生成ES数据"""


def auto_gen_es_file(file_wav_dict, wiki_token):
    sheet_id_list, _, excel_id = cloudfeishu_h.get_sheet_id_list(wiki_token)
    # print(sheet_id_list)
    # 查找要导入的媒体文件里有没有对应的
    for file_wav_name in file_wav_dict:
        flag = 0
        # 读飞书在线表
        for sheet_id in sheet_id_list:
            # 获取表格的标题列及表格数据
            title_colunmn_dict, values = cloudfeishu_h.get_sheet_title_column(wiki_token, sheet_id, excel_es_title_list)
            # print(title_colunmn_dict)
            # {'文件名': 2, 'External_Type': 8}

            # 遍历文件名列
            if '文件名' in title_colunmn_dict.keys():
                for row_index in range(1, len(values)):
                    column_letter = cloudfeishu_h.col_index_to_letter(title_colunmn_dict['文件名'])
                    # 读取测试
                    file_name = cloudfeishu_h.get_sheet_row_and_column_value(column_letter, row_index, excel_id,
                                                                             sheet_id)
                    if file_name in file_wav_name:
                        flag = 1
                        # print(file_wav_name)

        if flag == 0:
            oi_h.print_error(file_wav_name + "：在语音需求表中不存在，请检查名称是否正确或在表格中补充该名字")


with WaapiClient() as client:
    """自动生成externalsource"""


    def gen_external():
        args = {
            "sources": [
                {
                    "input": config.external_input_path,
                    "platform": "Windows",
                    "output": os.path.join(config.external_output_win_path, language)
                },
                {
                    "input": config.external_input_path,
                    "platform": "Android",
                    "output": os.path.join(config.external_output_android_path, language)
                },
                {
                    "input": config.external_input_path,
                    "platform": "iOS",
                    "output": os.path.join(config.external_output_ios_path, language)
                }
            ]
        }

        gen_log = client.call("ak.wwise.core.soundbank.convertExternalSources", args)


    # 查找External下的VO
    obj_list = client.call("ak.wwise.core.object.get",
                           waapi_h.waql_by_type("ExternalSource", config.vo_external_path),
                           options=config.options)['return']
    external_vo_list, _, _ = waapi_h.find_obj(obj_list)
    # pprint(external_sound_list)

    # 查找External下的CG
    obj_list = client.call("ak.wwise.core.object.get",
                           waapi_h.waql_by_type("ExternalSource", config.cg_external_path),
                           options=config.options)['return']
    external_cg_list, _, _ = waapi_h.find_obj(obj_list)
    # pprint(external_cg_list)

    """*******************主程序*******************"""

    # 语音ES生成
    for language in config.language_list:
        wav_language_path = os.path.join(py_path, "New_Media", language)
        file_wav_language_dict, _ = oi_h.get_type_file_name_and_path('.wav', wav_language_path)
        auto_gen_es_file(file_wav_language_dict, excel_es_vo_token)

    # 音效ES生成
    wav_path = os.path.join(py_path, "New_Media", "SFX")
    file_wav_dict, _ = oi_h.get_type_file_name_and_path('.wav', wav_path)

    # pprint(file_wav_dict)

    # # xml文件写入
    # with open('ExternalSource.xml', 'w+') as f:
    #     # 按照格式写入
    #     f.write(doc.toprettyxml())
    #     f.close()
    #
    # # 复制xml为wsources
    # source_path = shutil.copy2(os.path.join(py_path, config.es_xml_path),
    #                            os.path.join(py_path, config.es_wsources_path))
    #
    # # 自动生成externalsource
    # gen_external()
    #
    # # 删除xml的内容
    # doc = parse(config.es_xml_path)
    # parent_node = doc.getElementsByTagName('ExternalSourcesList')[0]
    # while parent_node.hasChildNodes():
    #     parent_node.removeChild(parent_node.firstChild)
    #
    # # 将删除的xml内容写入wsources
    # source_path = shutil.copy2(os.path.join(py_path, config.es_xml_path),
    #                            os.path.join(py_path, config.es_wsources_path))
    #
    # # 保存excel表的信息
    # wb_mediainfo.save(config.excel_mediainfo_path)
    # wb_wwisecookie.save(config.excel_wwisecookie_path)
    #
    # # 将excel转为csv
    # # 指定CSV文件名和编码格式
    # encoding = 'utf-8'
    # # media_info csv写入
    # df = pd.read_excel(config.excel_mediainfo_path)
    # df.to_csv(media_info_path, encoding=encoding, index=False)
    # # external_cookie csv写入
    # df = pd.read_excel(config.excel_wwisecookie_path)
    # df.to_csv(external_cookie_path, encoding=encoding, index=False)
    #
    # # 清除复制的媒体资源
    # shutil.rmtree("New_Media")
    # os.mkdir("New_Media")
    # os.mkdir("New_Media/Chinese")
    # os.mkdir("New_Media/English")
    # os.mkdir("New_Media/Japanese")
    # os.mkdir("New_Media/Korean")

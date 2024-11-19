from os.path import abspath, dirname
import sys
import os

"""获取py的根路径"""


def get_py_path():
    py_path = ""
    if hasattr(sys, 'frozen'):
        py_path = dirname(sys.executable)
    elif __file__:
        py_path = dirname(abspath(__file__))
    return py_path


"""获取主目录下的文件"""


def get_py_file_path(script_name):
    py_file_path = os.path.join(
        get_py_path(), "files", script_name)
    return py_file_path


"""Wwise"""
wwise_sfx_path = "\\Actor-Mixer Hierarchy\\v1"
wwise_event_path = "\\Events\\v1"
# 语音媒体资源的路径
wwise_vo_media_path = r'S:\chen.gong_DCC_Audio\Audio\SilverPalace_WwiseProject\Originals\Voices'
# Wwise工程路径
wwise_proj_path = r"S:\\chen.gong_DCC_Audio\\Audio\\SilverPalace_WwiseProject"

"""External Source"""
# excel表路径
excel_mediainfo_path = 'MediaInfoTable.xlsx'
excel_wwisecookie_path = 'ExternalSourceDefaultMedia.xlsx'
excel_DT_AudioPlotInfo_path = "DT_AudioPlotInfo.xlsx"
excel_DT_AudioPlotSoundInfo_path = "DT_AudioPlotSoundInfo.xlsx"

# excel表标题列
excel_es_title_list = ["文件名", "External_Type", "State", "剧情ID"]

# csv表路径
csv_mediainfo_path = "MediaInfoTable.csv"
csv_wwisecookie_path = "ExternalSourceDefaultMedia.csv"
csv_DT_AudioPlotInfo_path = "DT_AudioPlotInfo.csv"
csv_DT_AudioPlotSoundInfo_path = "DT_AudioPlotSoundInfo.csv"

# Wwise ES路径
vo_external_path = "\\Actor-Mixer Hierarchy\\v1\\VO\\VO\\VO_External"
cg_external_path = "\\Actor-Mixer Hierarchy\\v1\\CG\\CG\\CG_External"

# xml路径
es_xml_path = 'ExternalSource.xml'
es_wsources_path = 'ExternalSource.wsources'

# external的输入输出路径
external_input_path = "/module/waapi/waapi_Auto_Import_ExternalSource/ExternalSource.wsources"
external_output_win_path = 'S:\\Ver_1.0.0\\Project\\Content\\Audio\\GeneratedExternalSources\\Windows'
external_output_android_path = 'S:\\Ver_1.0.0\\Project\\Content\\Audio\\GeneratedExternalSources\\Android'
external_output_ios_path = 'S:\\Ver_1.0.0\\Project\\Content\\Audio\\GeneratedExternalSources\\iOS'
external_output_path = 'S:\\Ver_1.0.0\\Project\\Content\\Audio\\GeneratedExternalSources'
language_list = ['Chinese', 'English', 'Japanese', 'Korean', "SFX"]

es_id_config = {
    "B类（LevelSequence）剧情音效":
        {
            "min": 1,
            "max": 5000
        },
    "B类（LevelSequence）剧情语音":
        {
            "min": 50001,
            "max": 60000
        },
    "C类（操控演员表演，接管镜头）剧情音效":
        {
            "min": 5001,
            "max": 10000
        },
    "C类（操控演员表演，接管镜头）剧情语音":
        {
            "min": 60001,
            "max": 70000
        },
    "D类（操控演员表演，不接管镜头）剧情音效":
        {
            "min": 10001,
            "max": 15000
        },
    "D类（操控演员表演，不接管镜头）剧情语音":
        {
            "min": 70001,
            "max": 80000
        }
}

es_event_id_dict = {
    "CG_External_1P_2D": 36003,
    "VO_External_1P_2D": 15001
}

"""WAAPI查询"""
# 查询结果返回值
options = {
    'return': ['name', 'id', 'notes', 'originalWavFilePath', 'isIncluded', 'IsLoopingEnabled',
               'musicPlaylistRoot', 'LoopCount', 'PlaylistItemType', 'owner', 'parent', 'type',
               'Sequences', 'path', 'shortId', 'Inclusion', 'Target', 'ActionType']

}

"""cloudfeishu"""
# 需要先申请应用程序
app_id = "cli_a67dd70d527bd013"
app_secret = "kw5O5VboETIgZERiNe39ebqL72kzbruu"

# wiki token
# ES语音表token
excel_es_vo_token_list = ["PVNYwZbdVi1LMhkhvbVch4cJn4C"]
# ES音效表token
excel_es_sfx_token_list = ["AsW4wmAbQiIHeLkI2cscdN0Ln3e"]
# 音频资源命名配置token
check_name_token = "K2Cfwl44XiMe6Wkhi7AcJ0fGnVb"

"""DT Audio"""
excel_dcc_dt_audio_path = r"S:\chen.gong_DCC_Audio\Audio\Config\Audio.xlsx"
excel_dt_audio_path = "Audio.xlsx"
dt_audio_sheet_name = "audio"
csv_dt_audio_path = "Audio.csv"
# 自动生成
rowname_index = 1
audioid_index = 2
audioname_index = 3
audioevent_index = 4
desc_index = 5
isloop_index = 6

# 手动配置
FadeDuration_index = 7
FadeCurveNum_index = 8
ObjectType_index = 9

event_id_config = {
    "Amb":
        {
            "min": 1,
            "max": 3000
        },
    "Char":
        {
            "min": 3001,
            "max": 7000
        },
    "Imp":
        {
            "min": 7001,
            "max": 8000
        },
    "Mon":
        {
            "min": 8001,
            "max": 13000
        },
    "Mus":
        {
            "min": 13001,
            "max": 14000
        },
    "Sys":
        {
            "min": 14001,
            "max": 15000
        },
    "VO":
        {
            "min": 15001,
            "max": 30000
        },
    "Set_State":
        {
            "min": 30001,
            "max": 32000
        },
    "Set_Switch":
        {
            "min": 32001,
            "max": 34000
        },
    "Set_Trigger":
        {
            "min": 34001,
            "max": 36000
        },
    "CG":
        {
            "min": 36001,
            "max": 40000
        },

}

"""UE"""
ue_event_path = r"S:\Ver_1.0.0\Project\Content\Audio\WwiseAudio\Events\v1"
# GE路径
ue_ge_path = r'S:\Ver_1.0.0\Project\Content\SPSkill'

"""资源导入"""
# 导入规则json
check_name_json = '导入规则.json'
# 状态列表
status_list = ['占位', '临时资源', 'to do', 'in progress', 'done']
# 要删除的状态列表
del_status_list = ['占位', '临时资源', 'to do', 'in progress', 'done', 'cancel']

"""命名规范检查"""
excel_check_name = '命名规范检查表.xlsx'

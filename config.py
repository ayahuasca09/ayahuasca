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
# State的路径
wwise_state_path = '\\States\\Default Work Unit'
# State Bank的路径
wwise_state_bank_path = '\\SoundBanks\\Set\\Set_State'
# Event的实体路径
wwise_event_workunit_path = r"S:\chen.gong_DCC_Audio\Audio\SilverPalace_WwiseProject\Events"
# Bus根路径
wwise_bus_path = "\\Master-Mixer Hierarchy\\v1\\Master Audio Bus"
# Aux Ducking Bus路径
wwise_aux_duck_path = "\\Master-Mixer Hierarchy\\v1\\Master Audio Bus\\Auxiliary\\Aux_Ducking\\Aux_Ducking"
# State Music路径
# wwise_mus_state_path = "\\Interactive Music Hierarchy\\v1\\Mus"
# State Amb路径
# wwise_amb_state_path = "\\Actor-Mixer Hierarchy\\v1\\Amb\\Amb\\Amb_Global\\Amb_Global\\Amb_Global_Area"
# Mus_Global路径
wwise_mus_global_path = "\\Interactive Music Hierarchy\\v1\\Mus\\Mus_Global"
# Amb_Global路径
wwise_amb_global_path = "\\Actor-Mixer Hierarchy\\v1\\Amb\\Amb\\Amb_Global\\Amb_Global\\Amb_Global_Area\\Amb_Global_Area\\Amb_Global_Area"
"""state folder path"""
state_path = '{59A9E7B9-61C9-461A-A96F-4997B541A59C}'
"""music path"""
music_path = "{E8A5F6A0-AB3E-4166-BB68-C7D7ECC92B29}"
stin_path = "{DD48832A-AAC1-49F7-B57B-86F5FACD8FB3}"
trig_path = "{94BBA60A-29EF-4328-83E8-075770BF2687}"
trig_event_path = "{EDD34B59-3938-4C22-A63D-838FD8DCD57E}"
"""媒体资源路径"""
original_path = r'S:\chen.gong_DCC_Audio\Audio\SilverPalace_WwiseProject\Originals\SFX'
"""rtpc路径"""
wwise_rtpc_path = "\\Game Parameters"
"""meter_path"""
wwise_effect_ducking_path = "\\Effects\\Ducking"

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

# external的输入输出路径
external_input_path = r"F:\pppppy\SP\module\waapi\waapi_Auto_Import_ExternalSource\ExternalSource.wsources"
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
    "VO_External_1P_2D": 15001,
    "VO_External_1P_3D": 15157,
    "VO_External_3P_2D": 15158,
    "VO_External_3P_3D": 15159,
    "VO_External_NPC_2D": 15156,
    "VO_External_NPC_3D": 15155,

}

"""WAAPI查询"""
# 查询结果返回值
options = {
    'return': ['name', 'id', 'notes', 'originalWavFilePath', 'isIncluded', 'IsLoopingEnabled',
               'musicPlaylistRoot', 'LoopCount', 'PlaylistItemType', 'owner', 'parent', 'type',
               'Sequences', 'path', 'shortId', 'Inclusion', 'Target', 'ActionType', "Color", 'maxDurationSource',
               'ControlInput']

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
# 音效语音表
excel_media_token_list = ["TcbAwUoNriYh0Rk8gG0clcNCnTh", "CatewGc9miJplrkKVi2cvBd5nMb", "XXq2wK5dbiH7lWkbw36ciHK4nCc"]
# Wwise属性配置表token
property_config_token = "Uj2owpIpXiNCROkHUVxcp2H9nQc"
# State/Switch表token
state_sheet_token = "QE2ywINT9iLxS4kSTNTcXNK4nCh"
# Bus资源表token
bus_sheet_token = "F0ApwYVFViCQ6SkammRcs9cGnIh"
# Mus结构表token
mus_config_token = "OA69weFjaiNq13knGSVcNzV5n2d"
# Mus资源表token
mus_sheet_token_list = ["OcV8wP6cmi78GtkuR2UcM4bEnTe", "YDHHwNt61idv0SkhFMbcWNzXn4f", "OA69weFjaiNq13knGSVcNzV5n2d"]
# Amb资源表
amb_sheet_token = "AaZuw0vlWiyLCrkgxr6c6wqonmh"

"""ID表生成"""
excel_dcc_dt_audio_path = r"S:\chen.gong_DCC_Audio\Audio\Config\Audio.xlsx"
# 分页版dcc表
excel_dcc_dt_audio_page_path = r"S:\chen.gong_DCC_Audio\Audio\Config\Audio_Page.xlsx"
excel_dt_audio_path = "Audio.xlsx"
dt_audio_sheet_name = "audio"
csv_dt_audio_path = "Audio.csv"

# ID表中的长音效配置
long_sound_time = 5

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
HoldEventType_index = 10

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
    "Item":
        {
            "min": 40001,
            "max": 42000
        }

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
# 单个字符串长度限制
one_word_len = 10
# 根据_分段的word数限制
word_list_len = 7

"""媒体资源占位生成"""

"""DCC表的引用路径生成v2"""
# 剧情表路径
excel_plot_path = r'S:\Ver_1.0.0\Project\Designer\Excel\Scenario\ScenarioInstance'
# 大世界表路径
excel_world_path = r"S:\Ver_1.0.0\Project\Designer\Excel\WorldRegion\P_MapData.xlsx"
# 角色展示表路径
excel_show_path = r"S:\Ver_1.0.0\Project\Designer\Excel\Character\C_CharacterTable.xlsx"
# 战斗状态表路径
excel_combat_path = r"S:\Ver_1.0.0\Project\Designer\Excel\BattleField\C_BattleFieldSystem.xlsx"
# 3D点声源表路径
excel_3d_path = r"S:\Ver_1.0.0\Project\Designer\Excel\Audio\P_Audio3D.xlsx"
# 程序写死的dict
logic_id_refer_dict = {
    "13003": "包磊写死",
    "13001": "施博文写死"
}
# 天气系统dict
time_id_refer_dict = {
    "30033": "在UE找DT_SPTODMatrix",
    "30034": "在UE找DT_SPTODMatrix",
    "30035": "在UE找DT_SPTODMatrix",
    "30036": "在UE找DT_SPTODMatrix",
    "30037": "在UE找DT_SPTODMatrix",
    "30038": "在UE找DT_SPTODMatrix",
    "30039": "在UE找DT_SPTODMatrix",
    "30040": "在UE找DT_SPTODMatrix",
    "30041": "在UE找DT_SPTODMatrix",
    "30042": "在UE找DT_SPTODMatrix",
    "30043": "在UE找DT_SPTODMatrix",
    "30044": "在UE找DT_SPTODMatrix",

}

"""Wwise属性配置"""
# 属性配置表
excel_config_path = "Wwise属性配置表.xlsx"

"""State和Switch创建"""
# State/Switch表
excel_state_path = "State和Switch创建表.xlsx"
# 状态类通用后缀列表
list_state_suffix = ['State', 'Type']

"""Bus自动化"""
excel_bus_path = "Bus资源表.xlsx"
# 需要排除颜色设置的列表及以Aux开头的
bus_no_color_list = ['Audio Objects', 'Auxiliary', 'Main Mix', 'SFX', 'Passthrough Mix']

"""Mus自动化"""
excel_mus_config_path = "Mus结构表.xlsx"

"""Amb自动化"""
excel_amb_config_path = "全局环境音效表.xlsx"
sheet_amb_config = "Amb框架结构"
sheet_amb_media = "Amb"
# 结构与环境音效映射字典
amb_state_dict = {
    "Amb_Map": "Map_Type",
    "Amb_System": "System_Type"

}

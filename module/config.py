"""Wwise"""
wwise_sfx_path = "\\Actor-Mixer Hierarchy\\v1"
wwise_event_path = "\\Events\\v1"

"""External Source"""
# excel表路径
excel_mediainfo_path = 'MediaInfoTable.xlsx'
excel_wwisecookie_path = 'ExternalSourceDefaultMedia.xlsx'

# excel表标题列
excel_es_title_list = ["文件名", "External_Type", "State", "剧情ID"]

# csv表路径
csv_mediainfo_path = "MediaInfoTable.csv"
csv_wwisecookie_path = "ExternalSourceDefaultMedia.csv"

# Wwise ES路径
vo_external_path = "\\Actor-Mixer Hierarchy\\v1\\VO\\VO\\VO_External"
cg_external_path = "\\Actor-Mixer Hierarchy\\v1\\CG\\CG\\CG_External"

# xml路径
es_xml_path = 'ExternalSource.xml'
es_wsources_path = 'ExternalSource.wsources'

# external的输入输出路径
external_input_path = "F:\\pppppy\\SP\\module\\waapi\\waapi_Auto_Import_ExternalSource\\ExternalSource.wsources"
external_output_win_path = 'S:\\Ver_1.0.0\\Project\\Content\\Audio\\GeneratedExternalSources\\Windows'
external_output_android_path = 'S:\\Ver_1.0.0\\Project\\Content\\Audio\\GeneratedExternalSources\\Android'
external_output_ios_path = 'S:\\Ver_1.0.0\\Project\\Content\\Audio\\GeneratedExternalSources\\iOS'
external_output_path = 'S:\\Ver_1.0.0\\Project\\Content\\Audio\\GeneratedExternalSources'
language_list = ['Chinese', 'English', 'Japanese', 'Korean', "SFX"]

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

# ES语音表token
excel_es_vo_token_list = ["O90kwkWAdiDw9Lkv31GcZQwOnoh"]
# ES音效表token
excel_es_sfx_token_list = ["AsW4wmAbQiIHeLkI2cscdN0Ln3e"]

"""DT Audio"""
excel_dcc_dt_audio_path = r"S:\chen.gong_DCC_Audio\Audio\Config\Audio.xlsx"
excel_dt_audio_path = "Audio.xlsx"
dt_audio_sheet_name = "audio"
csv_dt_audio_path = "Audio.csv"
rowname_index = 1
audioid_index = 2
audioname_index = 3
audioevent_index = 4
desc_index = 5
isloop_index = 6

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

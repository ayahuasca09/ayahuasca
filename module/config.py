"""External Source"""
# excel表路径
excel_mediainfo_path = 'MediaInfoTable.xlsx'
excel_wwisecookie_path = 'ExternalSourceDefaultMedia.xlsx'

# excel表标题列
excel_es_title_list = ["文件名", "External_Type"]

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
language_list = ['Chinese', 'English', 'Japanese', 'Korean']

"""WAAPI查询"""
# 查询结果返回值
options = {
    'return': ['name', 'id', 'notes', 'originalWavFilePath', 'isIncluded', 'IsLoopingEnabled',
               'musicPlaylistRoot', 'LoopCount', 'PlaylistItemType', 'owner', 'parent', 'type',
               'Sequences', 'path', 'shortId']

}

"""cloudfeishu"""
# 需要先申请应用程序
app_id = "cli_a67dd70d527bd013"
app_secret = "kw5O5VboETIgZERiNe39ebqL72kzbruu"

# ES语音表token
excel_es_vo_token = "PVNYwZbdVi1LMhkhvbVch4cJn4C"
# ES音效表token
excel_es_sfx_token = "AsW4wmAbQiIHeLkI2cscdN0Ln3e"

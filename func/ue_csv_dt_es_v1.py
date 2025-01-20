import unreal
import sys
import os

# 获取当前脚本所在目录的父目录
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.abspath(os.path.join(current_dir, os.pardir))

# 将父目录添加到系统路径
sys.path.append(parent_dir)

# 现在可以导入 config_custom
import config_custom

# print(config_custom.wwise_path)


# csv_path = r"S:\chen.gong_DCC_Audio\Audio\Tool\Auto_Sound\ID表生成\Audio.csv"
csv_path_DT_AudioPlotInfo = os.path.join(current_dir, "DT_AudioPlotInfo.csv")
csv_path_DT_AudioPlotSoundInfo = os.path.join(current_dir, "DT_AudioPlotSoundInfo.csv")
csv_path_ExternalSourceDefaultMedia = os.path.join(current_dir, "ExternalSourceDefaultMedia.csv")
csv_path_MediaInfoTable = os.path.join(current_dir, "MediaInfoTable.csv")
# print(csv_path)
# ue_asset_path = "/Game/LogicRes/Audio/DT_AudioData"
ue_asset_DT_AudioPlotInfo = '/Game/LogicRes/Audio/AILanguage/DT_AudioPlotInfo'
ue_asset_ExternalSourceDefaultMedia = '/Game/LogicRes/Audio/AILanguage/DT_AudioExternalCookie'
ue_asset_MediaInfoTable = '/Game/LogicRes/Audio/AILanguage/DT_AudioMediaInfo'


def ue_csv_to_dt(ue_asset_path, csv_path):
    # Load DataTable asset
    data_table = unreal.EditorAssetLibrary.load_asset(ue_asset_path)

    if not data_table:
        unreal.log_error(f"无法从 {ue_asset_path} 加载数据表资产")
    else:
        try:
            # Fill in data from the CSV file
            unreal.DataTableFunctionLibrary.fill_data_table_from_csv_file(data_table, csv_path)
            # Save the updated DataTable asset
            unreal.EditorAssetLibrary.save_loaded_asset(data_table, True)
            unreal.log("数据表更新并成功保存。")
        except Exception as e:
            unreal.log_error(f"更新数据表失败：{e}")


ue_csv_to_dt(ue_asset_MediaInfoTable, csv_path_MediaInfoTable)
ue_csv_to_dt(ue_asset_ExternalSourceDefaultMedia, csv_path_ExternalSourceDefaultMedia)
# ue_csv_to_dt(ue_asset_DT_AudioPlotSoundInfo, csv_path_DT_AudioPlotSoundInfo)
ue_csv_to_dt(ue_asset_DT_AudioPlotInfo, csv_path_DT_AudioPlotInfo)

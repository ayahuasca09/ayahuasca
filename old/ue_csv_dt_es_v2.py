import unreal
import sys
import os
import re

# 获取当前脚本所在目录的父目录
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.abspath(os.path.join(current_dir, os.pardir))

# 将父目录添加到系统路径
sys.path.append(parent_dir)
import config_custom

# csv_path_DT_Merge = os.path.join(current_dir, "DT_ExternalSourceMediaAndCookieInfo.csv")
# ue_asset_Merge = '/Game/LogicRes/Audio/AILanguage/DT_ExternalSourceMediaAndCookieInfo'

# 存放ES表的文件夹
csv_excel_path = os.path.join(current_dir, "GenCSV")
ue_es_dt_path = '/Game/LogicRes/Audio/AudioPlotLanguage'

# UE的ES DT表路径
ue_path = config_custom.ue_path
es_dt_path = os.path.join(ue_path, "LogicRes", "Audio", "AudioPlotLanguage")

# ES DT表模板路径
ue_es_dt_temp_path = '/Game/LogicRes/FeatureTest/AudioTest/ESTemp'


# 检测资源路径是否存在，若不存在则创建路径
def ensure_directory_exists(directory_path):
    if not os.path.exists(directory_path):
        os.makedirs(directory_path)


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


# 检测字符串是否有中文，有则返回true
def contains_chinese(text):
    # 正则表达式用于匹配中文字符的范围
    pattern = re.compile(r'[\u4e00-\u9fff]')
    # 使用 search 方法查找字符串中是否有匹配的中文字符
    return pattern.search(text) is not None


# ue_csv_to_dt(ue_asset_Merge, csv_path_DT_Merge)

# 遍历所有csv，若未创建则创建，若已创建则导入

# 遍历目录下的 .csv 文件
# 遍历目录下的文件
for filename in os.listdir(csv_excel_path):
    csv_path = os.path.join(csv_excel_path, filename)
    # 检查文件是否以 .csv 结尾
    if filename.endswith('.csv'):
        if not contains_chinese(filename):
            ue_es_path_one = os.path.join(es_dt_path, filename.replace(".csv", ".uasset"))
            asset_path = ue_es_dt_path + "/" + filename.replace(".csv", "")
            # 若dt表不存在则创建
            if not os.path.exists(ue_es_path_one):
                unreal.EditorAssetLibrary.duplicate_asset(ue_es_dt_temp_path,
                                                          asset_path)
            # dt表导入
            ue_csv_to_dt(asset_path, csv_path)

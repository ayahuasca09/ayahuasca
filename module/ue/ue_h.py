import unreal


# csv_path = r"S:\chen.gong_DCC_Audio\Audio\Tool\Auto_Sound\ID表生成\Audio.csv"
# ue_asset_path = "/Game/LogicRes/Audio/DT_AudioData"


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

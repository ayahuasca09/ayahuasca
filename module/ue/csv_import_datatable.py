import unreal

fpath = r"S:\chen.gong_DCC_Audio\Audio\Tool\Auto_Sound\ID表生成\Audio.csv"  # Use csv to find the coresponse data table.
dpath = "/Game/LogicRes/Audio"  # only provide the folder


def my_reimport():
    task = unreal.AssetImportTask()
    task.filename = fpath
    task.destination_path = dpath
    task.replace_existing = True
    task.automated = True
    task.save = False

    csv_factory = unreal.CSVImportFactory()
    csv_factory.automated_import_settings.import_row_struct = unreal.load_object(None,
                                                                                 "/Script/CoreUObject.ScriptStruct/Script/SPGameFramework.SP_AudioConfigV1Set")  # use unreal.load_object to load struct type as an object.

    task.factory = csv_factory

    asset_tools = unreal.AssetToolsHelpers.get_asset_tools()
    asset_tools.import_asset_tasks([task])


my_reimport()

import json
from pprint import pprint


# wwise数据解析
def json_get_wwise_info(path):
    json_path = path
    with open(json_path, 'r') as jsfile:
        # 这是json转换后的dict，可在python里处理了
        js_dict = json.load(jsfile)
        # pprint(js_dict)
        js_IncludedEvents_list = []

        if 'SoundBanksInfo' in js_dict:
            js_SoundBanksInfo = js_dict['SoundBanksInfo']
            if 'SoundBanks' in js_SoundBanksInfo:
                js_SoundBanks_list = js_SoundBanksInfo['SoundBanks']
                js_SoundBanks_dict = js_SoundBanks_list[0]

                # Media文件解析
                if 'Media' in js_SoundBanks_dict:
                    js_Media_list = js_SoundBanks_dict['Media']
                    # # pprint(js_IncludedEvents_list)
                    for media_dict in js_Media_list:
                        media_path = media_dict['Path']
                        pprint("****" + media_path + "********")

        return js_IncludedEvents_list


"""查找特定名称的键"""


# json_data：python转化后的json数据
# value：特定键名
# key：特定键值


def find_specific_classname(json_data, value, key):
    # 使用列表推导式查找所有符合条件的字典
    return [item for item in json_data if item.get(key) == value]

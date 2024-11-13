from waapi import WaapiClient
import config as config
import module.waapi.waapi_h as waapi_h

args = {}
waql_type = ""
obj_id_or_path = ""
waql_id = ""

with WaapiClient() as client:
    """对象信息获取:通过类型获取所有子对象"""
    obj_list = client.call("ak.wwise.core.object.get",
                           waapi_h.waql_by_type(waql_type, obj_id_or_path),
                           options=config.options)['return']
    obj_list, obj_sub_id, obj_sub_path = waapi_h.find_obj(obj_list)

    # 通过类型查询所有
    obj_list = client.call("ak.wwise.core.object.get",
                           waapi_h.waql_from_type(waql_type),
                           options=config.options)['return']
    obj_list, obj_sub_id, obj_sub_path = waapi_h.find_obj(obj_list)

    # 通过ID查询对象
    obj_list = client.call("ak.wwise.core.object.get",
                           waapi_h.waql_from_type(waql_id),
                           options=config.options)['return']
    obj_list, obj_sub_id, obj_sub_path = waapi_h.find_obj(obj_list)

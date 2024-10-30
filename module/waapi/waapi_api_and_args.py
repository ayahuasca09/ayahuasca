from waapi import WaapiClient, CannotConnectToWaapiException
from pprint import pprint
import json
import module.config as config
import module.waapi.waapi_h as waapi_h

args = {}
waql_type = ""
obj_id_or_path = ""

with WaapiClient() as client:
    """对象信息获取"""
    obj_list = client.call("ak.wwise.core.object.get",
                           waapi_h.waql_by_type(waql_type, obj_id_or_path),
                           options=config.options)['return']
    obj_list, obj_sub_id, obj_sub_path = waapi_h.find_obj(obj_list)

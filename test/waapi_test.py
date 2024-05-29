from waapi import WaapiClient, CannotConnectToWaapiException
from pprint import pprint


with WaapiClient() as client:

    """基本功能测试"""
    args = {
        'waql': '$ "\\Events" select descendants where type = "Event" '
    }
    options = {
        'return': ["parent", "name"]
    }
    # 存储了所有Event的单元结构
    wwise_info_list = client.call("ak.wwise.core.object.get", args, options = options)['return']
    pprint(wwise_info_list)
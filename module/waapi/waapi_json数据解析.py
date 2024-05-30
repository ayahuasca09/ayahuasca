from waapi import WaapiClient, CannotConnectToWaapiException
from pprint import pprint
import config
import module.oi.oi_h as oi_h

file_name_json_list = oi_h.oi_get_json_filesname()
print(file_name_json_list)

with WaapiClient() as client:
    pass

from waapi import WaapiClient, CannotConnectToWaapiException
from pprint import pprint
import config
import module.oi.oi_h as oi_h

# file_json_list = oi_h.oi_get_json_filesname()
# pprint(file_json_list)

with WaapiClient() as client:
    client.call("ak.wwise.core.soundbank.generated")

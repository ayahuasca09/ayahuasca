from waapi import WaapiClient, CannotConnectToWaapiException
from pprint import pprint
import json

with WaapiClient() as client:
    result = client.call("ak.wwise.core.getProjectInfo")
    print(json.dumps(result, indent=4))

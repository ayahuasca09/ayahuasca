from waapi import WaapiClient, CannotConnectToWaapiException
from pprint import pprint

with WaapiClient() as client:
    args = {
        "soundbanks": [
            {"name": "ababa"}
        ],
        "writeToDisk": True
    }

    client.call("ak.wwise.core.soundbank.generate", args)

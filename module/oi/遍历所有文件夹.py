import os
from pprint import pprint


def list_all_files(directory):
    for root, dirs, files in os.walk(directory):
        for file in files:
            file_path = os.path.join(root, file)
            pprint(file_path)


# 使用示例
directory_path = r'S:\Ver_1.0.0\Project\Content\Audio\GeneratedSoundBanks'
list_all_files(directory_path)

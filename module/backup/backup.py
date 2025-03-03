import os
import zipfile
from datetime import datetime
import func.comlib.config as config
import func.comlib.cloudfeishu_h as cloudfeishu_h
from concurrent.futures import ThreadPoolExecutor

"""
将某个路径的文件压缩为zip，压缩名称为（路径下的目录或文件名）+日期（需要月和日），需要指定输出路径
需要在传入的参数路径下再加个子文件夹，子文件夹名称也是当前日期的这个名字，若子文件夹不存在则创建，若存在则把压缩的输出路径变成传入参数输出路径+日期名子文件夹
"""


def zip_directory_or_file(input_path):
    # 确保输入路径存在
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"The input path {input_path} does not exist.")
    # 判断输入路径是文件还是目录
    if os.path.isdir(input_path):
        base_name = os.path.basename(input_path.rstrip('/\\'))
    else:
        base_name = os.path.splitext(os.path.basename(input_path))[0]

    # 构造输出的zip文件名
    zip_filename = f"{base_name}_{date_str}.zip"
    output_path = os.path.join(date_subfolder_path, zip_filename)

    # 压缩文件或目录
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        if os.path.isdir(input_path):
            # 压缩整个目录
            for root, dirs, files in os.walk(input_path):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, input_path)
                    zipf.write(file_path, arcname)
        else:
            # 压缩单个文件
            zipf.write(input_path, os.path.basename(input_path))

    print(f"Compressed {input_path} into {output_path}")


"""备份路径"""
output_dir = r'D:\LE'
# 获取当前日期字符串
date_str = datetime.now().strftime("%m%d")

# 创建日期子文件夹路径
date_subfolder_path = os.path.join(output_dir, date_str)
# feishu表路径
excel_path = os.path.join(date_subfolder_path, "Excel")

# 如果日期子文件夹不存在，则创建
if not os.path.exists(date_subfolder_path):
    os.makedirs(date_subfolder_path)

# 如果日期子文件夹不存在，则创建
if not os.path.exists(excel_path):
    os.makedirs(excel_path)

# 示例用法
"""文件备份"""
# input_path_list = [r"S:\chen.gong_DCC_Audio",
#                    r"S:\Ver_1.0.0\Project\Content\Audio",
#                    r"S:\Ver_1.0.0\Project\Designer\Excel",
#                    r"S:\Ver_1.0.0\Project\Content\LogicRes"
#                    ]
#
# for input_path in input_path_list:
#     zip_directory_or_file(input_path)

"""feishudonwn"""
# # 合并字符串
# amb_sheet_dict = {"全局环境音效表": r"ZbeBwK43uiq55ekqZpccOq9NnCo"}
# down_dict = config.mus_sheet_token_dict | amb_sheet_dict | config.es_sheet_token_dict | config.media_sheet_token_dict
# tasks = []
# # 使用 ThreadPoolExecutor 来并行执行任务
# with ThreadPoolExecutor() as executor:
#     for media_excel_name in down_dict:
#         excel_token = down_dict[media_excel_name]  # 获取对应的表令牌
#         # 提交下载任务到线程池
#         # 第一个参数是函数名，后面参数是函数参数
#         task = executor.submit(cloudfeishu_h.download_media_excel, media_excel_name, excel_token, date_subfolder_path)
#         tasks.append(task)  # 将任务添加到任务列表中
# # 等待所有任务完成（可选）
# for task in tasks:
#     task.result()  # 调用 result() 确保任务完成，并捕获可能的异常

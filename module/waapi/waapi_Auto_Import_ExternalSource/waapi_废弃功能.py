"""读csv表"""
# with open(media_info_path, 'r+', newline='') as csvfile:
#     media_info_reader = csv.DictReader(csvfile)
#     media_info_writer = csv.DictWriter(csvfile, media_info_header)
#
#     flag = 0
#     for media_info_row in media_info_reader:
#         # print(media_info_row['Name'])
#         if vo_id_column:
#             if media_info_row['Name'] == sheet.cell(row=cell_sound.row,
#                                                     column=vo_id_column).value:
#                 media_info_row = {
#                     "Name": media_info_row['Name'],
#                     'ExternalSourceMediaInfoId': media_info_row['Name'],
#                     'MediaName': cell_sound.value + ".wem",
#                     'CodecID': 4,
#                     'bIsStreamed': 'TRUE',
#                     'bUseDeviceMemory': "FALSE",
#                     'MemoryAlignment': 0,
#                     'PrefetchSize': 0
#                 }
#                 # media_info_writer.writerow(new_media_info_row)
#                 flag = 1
#                 break
#     media_info_writer.writerows(media_info_reader)
#     # 未找到ID则添加新信息
#     if flag == 0:
#         new_media_info_row = {
#             "Name": sheet.cell(row=cell_sound.row,
#                                column=vo_id_column).value,
#             'ExternalSourceMediaInfoId': sheet.cell(row=cell_sound.row,
#                                                     column=vo_id_column).value,
#             'MediaName': cell_sound.value + ".wem",
#             'CodecID': 4,
#             'bIsStreamed': 'TRUE',
#             'bUseDeviceMemory': "FALSE",
#             'MemoryAlignment': 0,
#             'PrefetchSize': 0
#         }
#         media_info_writer.writerow(new_media_info_row)

"""读取csv"""

# filename：读取的csv文件路径
# row_name：要读取的列名

# def read_csv(filename, row_name):
#     for row in reader:
#         # print((row[row_name]))
#         pass


# with open(external_cookie, 'r') as csvfile:
#     external_cookie_reader = csv.DictReader(csvfile)

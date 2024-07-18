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


"""写入media_info csv表"""

# initial_row = media_info_data.shape[0]
# row_max = media_info_data.shape[0] + 1
# def write_media_info_csv():
#     global row_max
#     if initial_row == 0:
#         new_data = pd.DataFrame([{
#             "Name": row_max,
#             'ExternalSourceMediaInfoId': vo_id,
#             'MediaName': cell_sound.value + ".wem",
#             'CodecID': 4,
#             'bIsStreamed': 'TRUE',
#             'bUseDeviceMemory': "FALSE",
#             'MemoryAlignment': 0,
#             'PrefetchSize': 0
#         }])
#         new_data.to_csv(media_info_path, mode='a', header=False,
#                         index=False)
#         row_max = row_max + 1
#     else:
#         # 在media_info表中未找到该ID，则新增
#         if media_info_data[media_info_data['ExternalSourceMediaInfoId'] == vo_id].empty:
#
#             new_data2 = [row_max,
#                          vo_id,
#                          cell_sound.value + ".wem",
#                          4,
#                          'TRUE',
#                          "FALSE",
#                          0,
#                          0
#                          ]
#             media_info_data.loc[row_max] = new_data2
#             media_info_data.to_csv(media_info_path, mode='a', header=False,
#                                    index=False)
#             row_max = row_max + 1
#
#         # 在media_info表中找到该ID，则替换
#         else:
#             media_info_data.loc[
#                 media_info_data['ExternalSourceMediaInfoId'] == vo_id, ['MediaName', 'CodecID',
#                                                                         'bIsStreamed',
#                                                                         'bUseDeviceMemory',
#                                                                         'MemoryAlignment',
#                                                                         'PrefetchSize']] = [cell_sound.value + ".wem",
#                                                                                             4, 'TRUE', 'FALSE', 0, 0]
#             # media_info_data.loc[
#             #     media_info_data['ExternalSourceMediaInfoId'] == vo_id, 'CodecID'] = 4
#             # media_info_data.loc[
#             #     media_info_data['ExternalSourceMediaInfoId'] == vo_id, 'bIsStreamed'] = 'TRUE'
#             # media_info_data.loc[
#             #     media_info_data['ExternalSourceMediaInfoId'] == vo_id, 'bUseDeviceMemory'] = 'FALSE'
#             # media_info_data.loc[
#             #     media_info_data['ExternalSourceMediaInfoId'] == vo_id, 'MemoryAlignment'] = 0
#             # media_info_data.loc[
#             #     media_info_data['ExternalSourceMediaInfoId'] == vo_id, 'PrefetchSize'] = 0
#             media_info_data.to_csv(media_info_path, index=False)
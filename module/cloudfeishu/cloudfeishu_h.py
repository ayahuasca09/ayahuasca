import requests
import json
from pprint import pprint
import time
import module.config as config
import module.oi.oi_h as oi_h

url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal/"
sheets_base_url = "https://open.feishu.cn/open-apis/sheets/v2/spreadsheets"
access_id = "t-g104avfnX4I53QVTO2DHLZWRNXCT632VSRMC2QOY"

"""获取acess token"""


def get_access_id():
    post_data = {
        "app_id": config.app_id,
        "app_secret": config.app_secret
    }
    r = requests.post(url, data=post_data)
    tat = r.json()["tenant_access_token"]
    return tat


"""获取header"""


def get_header():
    header = {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + str(get_access_id())
    }  # 请求头
    return header


"""获取wiki的token"""


def get_excel_token(wiki_token):
    url = f"https://open.feishu.cn/open-apis/wiki/v2/spaces/get_node?token={wiki_token}"
    response = requests.get(url, headers=get_header())
    node_info = response.json()
    token = node_info["data"]["node"]["obj_token"]
    return token


"""获取表格ID"""


def get_sheet_id_list(wiki_token):
    excel_id = get_excel_token(wiki_token)
    url = f"https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/{excel_id}/sheets/query"

    response = requests.get(url, headers=get_header())
    sheets_info = response.json()
    sheet_id_list = []
    sheet_id_name_dict = {}
    # 是一整个列表的信息
    sheet_list = sheets_info["data"]["sheets"]
    for sheet in sheet_list:
        sheet_id_list.append(sheet['sheet_id'])
        sheet_id_name_dict[sheet['sheet_id']] = sheet['title']
    return sheet_id_list, sheet_id_name_dict, excel_id


"""获取表格所有内容"""


def get_sheet_all_content(wiki_token, sheet_id):
    sheet_all_url = f"{sheets_base_url}/{get_excel_token(wiki_token)}/values/{sheet_id}!A:Z"
    ret = requests.get(sheet_all_url, headers=get_header())
    # print(ret.text)
    data = ret.json()
    # pprint(data)
    return data


"""获取表格标题列"""


def get_sheet_title_column(wiki_token, sheet_id, title_list):
    data = get_sheet_all_content(wiki_token, sheet_id)
    # 标题所在列的字典映射
    title_colunmn_dict = {}
    # 如果响应成功code会返回0
    if data["code"] == 0:
        # 表格数据
        values = data["data"]["valueRange"]["values"]
        # 获取第一行的数据
        sheet_header = values[0]

        # 查找第一行中数据为xxx的，获取列索引column_index
        for i, value in enumerate(sheet_header):
            # print(value)
            for title_name in title_list:
                if value == title_name:
                    title_colunmn_dict[value] = i

    else:
        oi_h.print_error("读取" + sheet_id + "失败:" + data["msg"])
    return title_colunmn_dict, values


"""将列索引转换为字母表示形式"""


# 代码中的 update_range = f"Sheet1!{chr(65 + column_index)}{row_index + 1}" 可能会有问题，
# 因为 chr(65 + column_index) 只适用于列索引在 0 到 25 之间的情况，
# 即列在 A 到 Z 之间。如果列索引超过 25，则需要考虑多字母列（如 AA、AB 等）
def col_index_to_letter(index):
    letter = ""
    while index >= 0:
        letter = chr(index % 26 + 65) + letter
        index = index // 26 - 1
    return letter


"""获取某行某列的数据"""


def get_sheet_row_and_column_value(column_letter, row_index, excel_id, sheet_id):
    url = f"{sheets_base_url}/{excel_id}/values/{sheet_id}!{column_letter}{row_index + 1}:{column_letter}{row_index + 1}"
    ret = requests.get(url, headers=get_header())
    data = ret.json()
    # {'code': 0,
    #  'data': {'revision': 84,
    #           'spreadsheetToken': 'ZXAksWz0UhgLVGtvQ5fcH48qnrb',
    #           'valueRange': {'majorDimension': 'ROWS',
    #                          'range': '1FDwtz!C2:C2',
    #                          'revision': 84,
    #                          'values': [['VO_External_Emotion_C01_01']]}},
    #  'msg': 'success'}

    value = data['data']['valueRange']['values'][0][0]
    # print(value)
    return value


"""获取标题列下的某行的值"""


def get_title_row_and_column_value(title_name, title_colunmn_dict, row_index, excel_id, sheet_id):
    file_name = ""
    if title_name in title_colunmn_dict.keys():
        # 为防止过界
        column_letter = col_index_to_letter(title_colunmn_dict[title_name])
        # 读取测试
        file_name = get_sheet_row_and_column_value(column_letter, row_index,
                                                   excel_id,
                                                   sheet_id)
    else:
        oi_h.print_error(title_name + ":标题列不存在，请检查表格中是否有该标题的列表")
    return file_name


"""将在线表转list处理"""


# list保存的每一行为dict，dict为每行所需列的数据
# wiki_token_list：可以放多个云文档excel的id
# need_title_list:所需的标题列，会保存这些标题下的数据
def convert_sheet_to_list(wiki_token_list, need_title_list):
    # 将所需要的表格数据转为一个list，每一个dict存取每行所需的数据
    sheet_list = []
    for wiki_token in wiki_token_list:
        sheet_id_list, _, excel_id = get_sheet_id_list(wiki_token)

        # 读飞书在线表
        for sheet_id in sheet_id_list:
            # 获取表格的标题列及表格数据
            title_colunmn_dict, values = get_sheet_title_column(wiki_token, sheet_id, need_title_list)
            # print(title_colunmn_dict)
            # {'文件名': 2, 'External_Type': 8}

            # 遍历列
            for row_index in range(1, len(values)):
                # 每行所需要的数据
                every_row_content_dict = {}
                # 获取每行数据
                for title_name in need_title_list:
                    every_row_content_dict[title_name] = get_title_row_and_column_value(title_name,
                                                                                        title_colunmn_dict,
                                                                                        row_index,
                                                                                        excel_id,
                                                                                        sheet_id)
                # 若为cancel的则不加入
                if every_row_content_dict['State'] != 'cancel':
                    sheet_list.append(every_row_content_dict)

    return sheet_list


"""下载云文档"""


def download_cloud_sheet(wiki_token, sheet_save_path):
    access_token = get_access_id()

    # 创建导出任务
    url1 = "https://open.feishu.cn/open-apis/drive/v1/export_tasks"
    headers = {"Content-Type": "application/json", "Authorization": "Bearer " + str(access_token)}  # 请求头
    payload1 = json.dumps({
        "file_extension": "xlsx",
        "token": wiki_token,
        "type": "sheet"
    })
    response = requests.request("POST", url=url1, data=payload1, headers=headers)
    print("1:" + response.json()["msg"])
    ticket = response.json()["data"]["ticket"]
    print("2:" + ticket)
    time.sleep(5)
    # 查看导出任务结果
    url2 = "https://open.feishu.cn/open-apis/drive/v1/export_tasks/" + ticket + "?token=" + wiki_token
    headers2 = {'Authorization': 'Bearer ' + str(access_token)}  # 请求头
    response = requests.request("GET", url2, headers=headers2)
    file_token = response.json()["data"]["result"]["file_token"]
    # file_token=token
    # pprint(response.json())
    print("3:" + file_token)
    # 下载文件
    url = "https://open.feishu.cn/open-apis/drive/v1/export_tasks/file/" + file_token + "/download"
    payload = ''
    response = requests.request("GET", url, headers=headers, data=payload)
    if response.status_code == 200:
        with open(sheet_save_path, "wb") as f:
            f.write(response.content)
        print("下载成功")
    else:
        print(f"下载失败: {response.status_code}, {response.text}")

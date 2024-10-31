import requests
import json
from pprint import pprint
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

import requests
import json
from pprint import pprint

# 参考文档：https://blog.musnow.top/posts/3588260369/index.html
# 处理的飞书表格链接:https://rzvo5fieru.feishu.cn/wiki/AsW4wmAbQiIHeLkI2cscdN0Ln3e?sheet=1FDwtz
# 处理的飞书表格名称：ES音效管理表

url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal/"
sheets_base_url = "https://open.feishu.cn/open-apis/sheets/v2/spreadsheets"

# 获取acess token
post_data = {
    "app_id": "cli_a67dd70d527bd013",
    "app_secret": "kw5O5VboETIgZERiNe39ebqL72kzbruu"
}
r = requests.post(url, data=post_data)
tat = r.json()["tenant_access_token"]
# print(tat)

header = {
    "Content-Type": "application/json",
    "Authorization": "Bearer " + str(tat)
}  # 请求头

# 获取云文档wiki的token
node_token = "AsW4wmAbQiIHeLkI2cscdN0Ln3e"
url = f"https://open.feishu.cn/open-apis/wiki/v2/spaces/get_node?token={node_token}"
response = requests.get(url, headers=header)
node_info = response.json()
token = node_info["data"]["node"]["obj_token"]
# print(token)

excel_id = token
sheet_url = f"{sheets_base_url}/{excel_id}/values"  # 写入的sh开头的文档地址，其他不变
# ES音效表的ID
sheet_id = "1FDwtz"
# 查询的title名称
title_name = "文件名"

"""*******************功能函数*******************"""

"""数据读取"""


def get_sheet_info(column_letter, row_index):
    url = f"{sheets_base_url}/{excel_id}/values/{sheet_id}!{column_letter}{row_index + 1}:{column_letter}{row_index + 1}"
    ret = requests.get(url, headers=header)
    # print(ret.text)
    # pprint(ret.json())
    # {'code': 0,
    #  'data': {'revision': 56,
    #           'spreadsheetToken': 'My8ksChcfhJ4gJtFDLacdN3xnhg',
    #           'valueRange': {'majorDimension': 'ROWS',
    #                          'range': '1FDwtz!C2:C2',
    #                          'revision': 56,
    #                          'values': [['Hello']]}},
    #  'msg': 'success'}


"""数据写入"""


# 传入的是列索引和行索引
def insert_sheet_info(column_letter, row_index, sheet_value):
    # range参数中!之前的工作簿ID，后面跟着行列号范围
    post_data = {
        "valueRange": {
            "range": f"{sheet_id}!{column_letter}{row_index + 1}:{column_letter}{row_index + 1}",
            "values": [[sheet_value]]
        }
    }
    # 在486268这个工作簿内的单元格C3到N8写入内容为helloworld等内容
    r2 = requests.put(sheet_url, data=json.dumps(post_data), headers=header)  # 请求写入
    # print(r2.json())  # 输出来判断写入是否成功
    update_result = r2.json()
    if update_result["code"] == 0:
        print("写入成功:", update_result['data']['updatedRange'], sheet_value)
    else:
        print("写入失败:", update_result["msg"], update_result["code"])


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


"""*******************主函数*******************"""

# 表格所有数据读取
sheet_all_url = f"{sheets_base_url}/{excel_id}/values/{sheet_id}!A:Z"
ret = requests.get(sheet_all_url, headers=header)
# print(ret.text)
data = ret.json()

# 如果响应成功code会返回0
if data["code"] == 0:
    # 表格数据
    values = data["data"]["valueRange"]["values"]
    # 获取第一行的数据
    sheet_header = values[0]
    column_index = None
    # 查找第一行中数据为xxx的，获取列索引column_index
    for i, value in enumerate(sheet_header):
        if value == title_name:
            column_index = i
            break

    if column_index is not None:
        # 遍历该列的数据并写入值
        for row_index in range(1, len(values)):
            column_letter = col_index_to_letter(column_index)
            # 写入测试
            # insert_sheet_info(column_letter, row_index, "Hello")
            # 读取测试
            # get_sheet_info(column_letter, row_index)
    else:
        print("未找到" + title_name)
else:
    print("读取" + sheet_id + "失败:", data["msg"])

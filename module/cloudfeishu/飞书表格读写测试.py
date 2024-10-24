import requests
import json
from pprint import pprint

# 参考文档：https://blog.musnow.top/posts/3588260369/index.html

url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal/"
sheets_base_url = "https://open.feishu.cn/open-apis/sheets/v2/spreadsheets"

# 获取acess token
post_data = {
    "app_id": "cli_a790b46698679013",
    "app_secret": "KbQp6PLKnIMgUBywe0lTIcmso1jw7abl"
}
r = requests.post(url, data=post_data)
tat = r.json()["tenant_access_token"]
# print(tat)

header = {
    "Content-Type": "application/json",
    "Authorization": "Bearer " + str(tat)
}  # 请求头

# 获取云文档wiki的token
node_token = "DSuEwyfFzi5jPgkc0owcTpgVnZc"
url = f"https://open.feishu.cn/open-apis/wiki/v2/spaces/get_node?token={node_token}"
response = requests.get(url, headers=header)
node_info = response.json()
token = node_info["data"]["node"]["obj_token"]
# print(token)

excel_id = token


# 数据写入
def insert_info():
    url = f"{sheets_base_url}/{excel_id}/values"  # 写入的sh开头的文档地址，其他不变
    # range参数中!之前的工作簿ID，后面跟着行列号范围
    post_data = {
        "valueRange": {
            "range": "vjLK9g!C3:D4",
            "values": [["Hello", 1], ["World", 1]]
        }
    }
    # 在486268这个工作簿内的单元格C3到N8写入内容为helloworld等内容
    r2 = requests.put(url, data=json.dumps(post_data), headers=header)  # 请求写入
    print(r2.json())  # 输出来判断写入是否成功


# insert_info()

# 数据读取
url = f"{sheets_base_url}/{excel_id}/values/vjLK9g!C1:C"
ret = requests.get(url, headers=header)
# print(ret.text)
pprint(ret.json())

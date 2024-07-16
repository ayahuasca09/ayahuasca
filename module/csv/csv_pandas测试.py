import pandas as pd

data = pd.read_csv('MediaInfoTable.csv')
# 显示数据的前五行
"""print(data.head())"""

# 显示数据的前10行
"""print(data[:10])"""

# 保留指定列数据
"""col = ['Name', 'PrefetchSize']
data = data[col]
print(data)"""

# 筛选指定行及多条件查询(ans)
"""data = data[(data['PrefetchSize'] == 0) & (data['Name'] == 200)]
print(data)"""

# 替换指定数据
data[data['Name'] == 200] = [201, 201, "New.wem", 2, "TRUE", "FALSE", 22, 0]
data.to_csv('MediaInfoTable.csv', index=False)
# print(type(ft_data))

# 新加数据
new_data = pd.DataFrame({
    "Name": 333,
    'ExternalSourceMediaInfoId': 333,
    'MediaName': ".wem",
    'CodecID': 4,
    'bIsStreamed': 'TRUE',
    'bUseDeviceMemory': "FALSE",
    'MemoryAlignment': 0,
    'PrefetchSize': 0
}, index=[0])
# index无效，但不加会报错
new_data.to_csv('MediaInfoTable.csv', mode='a', header=False, index=False)
# mode：不指定时如果csv文件不存在则自动创建

# 判断当表格中不存在此值
if data[data['Name'] == 333].empty:
    print(42352)


import pandas as pd

"""将excel文件写入csv"""


def excel_to_csv(excel_path, csv_media_path):
    encoding = 'utf-8'
    # media_info csv写入
    df = pd.read_excel(excel_path)
    # 将所有数字列转换为整数
    for col in df.select_dtypes(include=['number']).columns:
        df[col] = df[col].fillna(0).astype(int)
    df.to_csv(csv_media_path, encoding=encoding, index=False)

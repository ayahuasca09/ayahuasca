import pandas as pd

"""将excel文件写入csv"""


def excel_to_csv(excel_path, csv_media_path):
    encoding = 'utf-8'
    # media_info csv写入
    df = pd.read_excel(excel_path)
    df.to_csv(csv_media_path, encoding=encoding, index=False)

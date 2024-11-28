import os
import pandas as pd
from bs4 import BeautifulSoup

# 定義資料夾名稱
input_folder = 'file'  # 存放 .xls 的資料夾
output_folder = 'converted_csv'  # 存放轉換後 .csv 的資料夾

# 確保輸出資料夾存在
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# 迭代處理資料夾中的檔案
for input_file in os.listdir(input_folder):
    # 只處理 .xls 的檔案
    if input_file.endswith('.xls'):
        file_path = os.path.join(input_folder, input_file)

        # 檢查檔案是否為空
        if os.stat(file_path).st_size == 0:
            print(f"檔案 {input_file} 是空的，跳過處理。")
            continue

        # 解析 HTML
        print(f"開始解析檔案：{input_file}")
        try:
            # 讀取檔案內容
            with open(file_path, 'r', encoding='utf-8') as f:
                html_content = f.read()

            # 使用 BeautifulSoup 解析 HTML
            soup = BeautifulSoup(html_content, 'html.parser')
            table = soup.find('table')  # 查找 HTML 中的表格

            if table:
                # 將 HTML 表格轉換為 Pandas DataFrame
                df = pd.read_html(str(table))[0]

                # 設定輸出的檔案路徑
                output_file = os.path.join(output_folder, input_file.replace('.xls', '.csv'))

                # 儲存轉換結果到輸出資料夾
                df.to_csv(output_file, index=False, encoding='utf-8-sig')
                print(f"成功將檔案 {input_file} 轉換為 CSV：{output_file}")
            else:
                print(f"檔案 {input_file} 中未找到 HTML 表格，跳過處理。")

        except Exception as e:
            print(f"解析檔案 {input_file} 時發生錯誤：{e}")

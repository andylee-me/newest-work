import os
import pandas as pd
from bs4 import BeautifulSoup

# 定義檔案路徑
input_folder = 'converted_csv'  # 存放 CSV 的資料夾

# 遍歷資料夾中的所有檔案
for input_file in os.listdir(input_folder):
    if input_file.endswith('.csv'):  # 只處理 CSV 檔案
        # 原始檔案的路徑
        csv_path = os.path.join(input_folder, input_file)
        
        # 讀取原始 CSV 檔案
        try:
            raw_data = pd.read_csv(csv_path, encoding='utf-8')
        except Exception as e:
            print(f"無法讀取檔案 {input_file}，錯誤：{e}")
            continue

        # 假設 HTML 表格內容存放在第一列第一欄
        try:
            html_content = raw_data.iloc[0, 0]  # 修改為正確的位置
        except Exception as e:
            print(f"無法提取 HTML 內容，檔案：{input_file}，錯誤：{e}")
            continue

        # 使用 BeautifulSoup 解析 HTML
        soup = BeautifulSoup(html_content, 'html.parser')

        # 提取表格
        table = soup.find('table')
        if table:
            try:
                # 將 HTML 表格轉換為 Pandas DataFrame
                df = pd.read_html(str(table))[0]

                # 覆蓋儲存至同一個檔案
                df.to_csv(csv_path, index=False, encoding='utf-8-sig')

                print(f"成功解析並覆蓋儲存檔案：{input_file}")
            except Exception as e:
                print(f"解析表格錯誤，檔案：{input_file}，錯誤：{e}")
        else:
            print(f"未找到 HTML 表格內容，檔案：{input_file}")

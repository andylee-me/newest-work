import os
import pandas as pd
from bs4 import BeautifulSoup

# 定義檔案路徑
input_folder = 'converted_csv'  # 存放 CSV 的資料夾

for input_file in os.listdir(input_folder):
    if input_file.endswith('.csv'):  # 只處理 CSV 檔案
        csv_path = os.path.join(input_folder, input_file)
        
        # 檢查檔案是否為空
        if os.stat(csv_path).st_size == 0:
            print(f"檔案 {input_file} 是空的，跳過處理。")
            continue

        try:
            raw_data = pd.read_csv(csv_path, encoding='utf-8')
        except Exception as e:
            print(f"無法讀取檔案 {input_file}，錯誤：{e}")
            continue
        
        # 檢查資料是否為空
        if raw_data.empty:
            print(f"檔案 {input_file} 沒有任何資料，跳過處理。")
            continue
        
        # 嘗試提取 HTML 內容
        try:
            html_content = raw_data.iloc[0, 0]  # 修改為正確的位置或欄位名稱
        except IndexError as e:
            print(f"無法提取 HTML 內容，檔案：{input_file}，錯誤：{e}")
            continue

        # 使用 BeautifulSoup 解析 HTML
        soup = BeautifulSoup(html_content, 'html.parser')

        # 提取表格
        table = soup.find('table')
        if table:
            try:
                df = pd.read_html(str(table))[0]
                df.to_csv(csv_path, index=False, encoding='utf-8-sig')  # 覆蓋儲存
                print(f"成功解析並覆蓋儲存檔案：{input_file}")
            except Exception as e:
                print(f"解析表格錯誤，檔案：{input_file}，錯誤：{e}")
        else:
            print(f"未找到 HTML 表格內容，檔案：{input_file}")

import os
import pandas as pd

# 定義檔案夾路徑
folder_path = 'file'  # 儲存 XLS 檔案的資料夾
output_folder = os.path.join(os.getcwd(), 'converted_csv')  # 確保使用相對路徑

# 確保輸出資料夾存在
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# 遍歷資料夾中的檔案
for file_name in os.listdir(folder_path):
    file_path = os.path.join(folder_path, file_name)
    if file_name.endswith('.xls') or file_name.endswith('.xlsx'):  # 找到所有 XLS/XLSX 檔案
        try:
            # 嘗試用 Excel 格式讀取
            data = pd.read_excel(file_path, engine='xlrd')
            print(f"成功讀取 Excel 檔案：{file_name}")
        except Exception as e:
            print(f"{file_name} 可能不是標準 Excel 格式，錯誤：{e}")
            try:
                # 嘗試用 CSV 格式讀取
                data = pd.read_csv(file_path, encoding='utf-8')
                print(f"{file_name} 已用 CSV 格式成功讀取。")
            except Exception as e2:
                print(f"{file_name} 無法讀取，跳過。錯誤：{e2}")
                continue

        # 儲存為 CSV
        csv_file_name = file_name.replace('.xls', '.csv').replace('.xlsx', '.csv')
        csv_file_path = os.path.join(output_folder, csv_file_name)
        data.to_csv(csv_file_path, index=False, encoding='utf-8-sig')  # 儲存為 CSV，使用 UTF-8 編碼
        print(f"成功將 {file_name} 轉換為 {csv_file_name}")

import os
import pandas as pd

# 資料夾路徑
converted_csv_dir = "converted_csv"
codes_dir = "codes"

# 讀取對應的檔名映射表
rename_mapping_file = os.path.join(codes_dir, "rename_mapping.csv")
try:
    mapping_df = pd.read_csv(rename_mapping_file)
    print("成功讀取檔名映射表：")
    print(mapping_df)
except Exception as e:
    print(f"讀取檔名映射表失敗：{e}")
    exit()

# 確保映射表包含所需的欄位
if "original_filename" not in mapping_df.columns or "new_filename" not in mapping_df.columns:
    print("映射表必須包含 'original_filename' 和 'new_filename' 欄位")
    exit()

# 重新命名檔案
for index, row in mapping_df.iterrows():
    original_file = os.path.join(converted_csv_dir, row["original_filename"])
    new_file = os.path.join(converted_csv_dir, row["new_filename"])

    # 檢查原始檔案是否存在
    if os.path.exists(original_file):
        try:
            os.rename(original_file, new_file)
            print(f"檔案重新命名：{row['original_filename']} -> {row['new_filename']}")
        except Exception as e:
            print(f"重新命名失敗：{row['original_filename']} -> {row['new_filename']}，錯誤：{e}")
    else:
        print(f"原始檔案不存在：{row['original_filename']}")

print("檔案重新命名完成。")

import os

base_dir = os.getcwd()
print("base_dir: ", base_dir)

test_file_dir = os.path.join(base_dir, "test_file")
print("test_file_dir: ", test_file_dir)

import zipfile

# 設定目標資料夾
hw_folder = os.path.join(base_dir, "HW_folder")

# 取得 HW_folder 內的所有檔案
files = os.listdir(hw_folder)

# 篩選出壓縮檔（假設 HW_folder 內只會有一個壓縮檔）
zip_files = [f for f in files if f.lower().endswith(".zip")]

# 確保 HW_folder 內有且只有一個壓縮檔
if len(zip_files) == 1:
    zip_path = os.path.join(hw_folder, zip_files[0])  # 取得完整路徑
    
    # 解壓縮
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(hw_folder)  # 解壓縮到 HW_folder

    # 刪除壓縮檔
    os.remove(zip_path)
    print(f"已解壓縮 {zip_files[0]} 並刪除原壓縮檔")
    
elif len(zip_files) == 0:
    print("錯誤：HW_folder 內沒有找到壓縮檔")
else:
    print("錯誤：HW_folder 內有多個壓縮檔，請手動確認")


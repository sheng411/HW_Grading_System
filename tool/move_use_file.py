import os
import shutil


file_extension='.cpp'

def copy_header_files_to_students(base_dir):
    """
    將 test_file 資料夾內的所有 .h 檔案，複製到 HW_folder 中每個學生的資料夾裡。
    
    :param base_dir: 專案根目錄（例如：0hw_check 的絕對或相對路徑）
    """
    test_file_path = os.path.join(base_dir, "test_file")
    hw_folder_path = os.path.join(base_dir, "HW_folder")

    # 取得 test_file 中所有 .h 檔案
    h_files = [f for f in os.listdir(test_file_path) if f.endswith(file_extension)]

    # 確保找到學生的資料夾
    if not os.path.exists(hw_folder_path):
        print(f"找不到 HW_folder：{hw_folder_path}")
        return

    student_dirs = [d for d in os.listdir(hw_folder_path)
                    if os.path.isdir(os.path.join(hw_folder_path, d))]

    for student in student_dirs:
        student_path = os.path.join(hw_folder_path, student)
        for h_file in h_files:
            src = os.path.join(test_file_path, h_file)
            dst = os.path.join(student_path, h_file)
            shutil.copy2(src, dst)  # 使用 copy2 保留檔案的 metadata
            print(f"已複製 {h_file} 到 {student_path}")

    print(f"所有 {file_extension} 檔案皆已複製完成。")

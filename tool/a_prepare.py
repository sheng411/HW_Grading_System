import csv
import os
import shutil
import zipfile


#預設
update_dir = "update_zip"


def find_csv_file():
    csv_files = [f for f in os.listdir('.') if f.lower().endswith('.csv')]

    if not csv_files:
        print("當前目錄下沒有任何 CSV 檔案")
        exit()

    if len(csv_files) == 1:
        print(f"找到 CSV 檔案：{csv_files[0]}")
        return csv_files[0]

    print("找到多個 CSV 檔案")
    for idx, file in enumerate(csv_files, 1):
        print(f"{idx}. {file}")

    while True:
        try:
            choice = int(input("請選擇執行檔案(輸入編號):"))
            if 1 <= choice <= len(csv_files):
                return csv_files[choice - 1]
            else:
                print("請輸入有效的編號")
        except ValueError:
            print("請輸入數字")


def clean_csv(input_file: str, output_file: str):
    with open(input_file, newline='', encoding='utf-8-sig') as infile, \
         open(output_file, 'w', newline='', encoding='utf-8-sig') as outfile:

        reader = csv.reader(infile)
        writer = csv.writer(outfile)

        next(reader)  # 跳過欄位名稱列

        for row in reader:
            if row and row[0].startswith('參與者'):
                row[0] = row[0].replace('參與者', '').strip()
            writer.writerow(row)


def process_csv(output_csv,update_dir_path,pdf_dir_path):
    if not os.path.exists(update_dir_path):
        os.makedirs(update_dir_path)

    with open(output_csv, newline='', encoding='utf-8-sig') as csvfile:
        rows = csv.reader(csvfile)
        for row in rows:
            #print(row)
            dir_path = os.path.join(update_dir_path, 'Participant_' + row[0] + '_assignsubmission_file_')
            if not os.path.exists(dir_path):
                os.makedirs(dir_path)

            file_name = row[2] + '.pdf'
            file_path = os.path.join(pdf_dir_path, file_name)
            copy_path = os.path.join(dir_path, file_name)

            try:
                shutil.copy(file_path, copy_path)
                print(f"成功複製 {file_name}")
            except FileNotFoundError:
                print(f"找不到檔案：{file_name}")


def zip_participant_folders(zip_name,update_dir_path,update_dir_all_path):
    update_dir_all_path=os.path.join(update_dir_all_path,zip_name)
    with zipfile.ZipFile(update_dir_all_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for folder in os.listdir(update_dir_path):
            full_folder_path = os.path.join(update_dir_path, folder)
            if os.path.isdir(full_folder_path) and folder.startswith('Participant_'):
                for root, dirs, files in os.walk(full_folder_path):
                    for file in files:
                        full_path = os.path.join(root, file)
                        arcname = os.path.relpath(full_path, start=update_dir_path)
                        zipf.write(full_path, arcname)
    print(f"\n\n*已壓縮完成：{zip_name}*")


def run_prepare(hw_dir_path,file_name,zip_name,pdf_dir_path,update_dir_all_path):
    input_csv = find_csv_file()
    clean_csv(input_csv, file_name)
    print(f"結果輸出至：{file_name}\n\n")
    update_dir_path=os.path.join(hw_dir_path,update_dir)

    process_csv(file_name,update_dir_path,pdf_dir_path)
    zip_participant_folders(zip_name,update_dir_path,update_dir_all_path)


if __name__ == '__main__':
    base_dir = os.getcwd()
    output_name = "output123.csv"
    zip_name = "output123.zip"
    pdf_dir_path = os.path.join(base_dir,"0All_pdf")
    run_prepare(base_dir,output_name,zip_name,pdf_dir_path,base_dir)

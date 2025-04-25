import json
import os

json_file = "config.json"

def create_json_file():    
    # 讓使用者輸入各個欄位的值
    unzip = input("是否解壓縮(y/n): ") or "y"
    copy_file2student= input("是否拷貝測試檔案至學生資料夾(預設為n): ") or "n"
    file_extension = input("請輸入檔案副檔名(eg. .cpp): ")
    num_problems = int(input("請輸入 num_problems: "))

    # 讓使用者輸入 score 陣列
    score = []
    print("請輸入 score，每次輸入一個數字")

    for i in range(num_problems):
        #print(f"第 {i + 1} 題: ")
        score.append(int(input(f"第 {i + 1} 題配分: ")))

    print(f"\n總分: {sum(score)}")
    print(f"score: {score}\n")

    score_check=input("分數是否正確(預設為y): ")or "y"
    while score_check.lower() == "n":
        if score_check.lower() == "n":
            score_check_count=int(input("請重新輸入第幾題: "))
            if score_check_count>num_problems or score_check_count<=0:
                print(f"\n請輸入正確的題號!\n")
                continue
            score[score_check_count-1]=int(input(f"第 {score_check_count} 題配分: "))
            print(f"\n總分: {sum(score)}")
            print(f"score: {score}\n")
            score_check=input("分數是否正確(預設為y): ")or "y"
    
    cpp_file_two = input("是否有兩個 cpp 檔案(預設為n): ") or "n"
    if cpp_file_two.lower() == "y":
        cpp_file2_name = input("請輸入第二個 cpp 檔案名稱(eg. main.cpp): ")
    else:
        cpp_file2_name = ""    
    ctf_count = input("是否檢查題目答案檔案(預設為y): ")or "y"
    selection = input("作業編號(eg.02261): ")
    use_zip = input("是否將輸出壓縮成 zip (預設為n): ")or "n"

    return unzip, copy_file2student, file_extension, num_problems, score, ctf_count, selection, cpp_file2_name, use_zip

def add_json_file(unzip, copy_file2student, file_extension, num_problems, score, ctf_count, selection, cpp_file2_name, use_zip):
    # 建立 JSON 結構
    code_path = input("輸入檔案存放路徑[存放於/test_file中](預設為當前): ").strip() or os.getcwd()
    code_path = os.path.join(code_path, "test_file")
    if not os.path.exists(code_path):
        os.makedirs(code_path)  # 如果路徑不存在，則創建它
    json_path = os.path.join(code_path, json_file)

    data = {
        "unzip": unzip,
        "copy_file2student": copy_file2student,
        "file_extension": file_extension,
        "num_problems": num_problems,
        "score": score,
        "ctf_count": ctf_count,
        "selection": selection,
        "cpp_file2_name": cpp_file2_name,
        "use_zip": use_zip
    }
    print(f"\n\n{data}\n\n")

    # 將資料寫入 JSON 檔案
    with open(json_path, "w", encoding="utf-8") as file:
        json.dump(data, file, indent=2, ensure_ascii=False)

    print(f"已成功寫入 {json_file}")


if __name__ == "__main__":
    unzip, copy_file2student, file_extension, num_problems, score, ctf_count, selection, cpp_file2_name, use_zip=create_json_file()
    add_json_file(unzip, copy_file2student, file_extension, num_problems, score, ctf_count, selection, cpp_file2_name, use_zip)
    input("執行完畢，請按 Enter 鍵退出...")
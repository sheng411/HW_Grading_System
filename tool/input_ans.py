import os
import subprocess

compile_path = r"C:\msys64\mingw64\bin\g++.exe"

def main():
    problem = int(input("題號(預設為0): ").strip()or 0)
    num_tests = int(input("測資數量(預設為10): ")or 10)
    code_path = input("輸入測試程式路徑(預設為當前): ").strip() or os.getcwd()
    cpp_file_two = input("是否有兩個以上.cpp檔案需要編譯(預設為n): ") or "n"
    cpp_file_other_name = []
    if cpp_file_two.lower() == "y":
        cfn_count = int(input("請輸入有幾個額外的 .cpp 檔案需要編譯(eg. 2): "))
        for i in range(cfn_count):
            cpp_file_other_name.append(input(f"[{i+1}/{cfn_count}] 請輸入其他 .cpp 檔案名稱(eg. main1.cpp，僅輸入 mian 即可): "))
    else:
        pass

    path_check = input("測試程式路徑是否和檔案存放路徑相同?(預設為y): ") or "y"
    if path_check.lower() == "y":
        file_path = code_path
    else:
        file_path = input("輸入檔案存放路徑(預設為當前): ").strip() or os.getcwd()
    
    code_path = os.path.join(code_path, "sample_code")
    file_path = os.path.join(file_path, "test_file")
    if not os.path.exists(code_path):
        os.makedirs(code_path)  # 如果路徑不存在，則創建它
        return
    if not os.path.exists(file_path):
        os.makedirs(file_path)

    print(f"\n\n測試檔案路徑: {code_path}")
    print(f"檔案存放路徑: {file_path}")
    

    # 設定輸入與答案檔案名稱，例如 1input.txt 與 1ans.txt
    input_filename = f"{problem}input.txt"
    ans_filename = f"{problem}ans.txt"
    cpp_filename = f"{problem}.cpp"
    exe_filename = f"{problem}"

    input_path = os.path.join(file_path, input_filename)
    ans_path = os.path.join(file_path, ans_filename)
    cpp_path = os.path.join(code_path, cpp_filename)
    exe_path = os.path.join(code_path, exe_filename)


    # 初始化編譯命令列表
    compile_cmd = [compile_path]

    # 學生的 .cpp 檔（如 1.cpp）
    if os.path.isfile(cpp_path):
        compile_cmd.append(cpp_path)
    else:
         msg = f"*找不到 {cpp_filename},Pass {i}"
         print(f"{msg}")
         return

    # 其他 .cpp 檔案
    if cpp_file_other_name:
        for ocf in cpp_file_other_name:
            o_cpp_filename = f"{ocf}{problem}.cpp"
            o_cpp_path = os.path.join(code_path, o_cpp_filename)
            if os.path.isfile(o_cpp_path):
                compile_cmd.append(o_cpp_path)

    # 添加輸出選項
    compile_cmd.extend(["-o", exe_path])


    result = subprocess.run(compile_cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if result.returncode != 0:
        print(f"\n編譯 {os.path.basename(cpp_path)} 時發生錯誤: {result.stderr}")
        return
    else:
        print(f"\n成功編譯 {os.path.basename(cpp_path)} 為 {os.path.basename(exe_path)}.exe")
    

    print("\n請依序輸入測試資料(每組測試資料輸入完後，請連續按兩次 Enter 來結束輸入): ")

    with open(input_path, "w", encoding="utf-8") as inf, open(ans_path, "w", encoding="utf-8") as anf:
        for i in range(1, num_tests+1):
            print(f"\n\n第 {i} 組測試資料: ")
            test_data = ""
            lines = []
            
            while True:
                line = input()
                if line == "":
                    break
                    
                lines.append(line)
            
            # 移除最後一個空行（因為它只是用來結束輸入的）
            if lines and lines[-1] == "":
                lines = lines[:-1]
                
            # 將所有行轉換為測試資料
            test_data = "\n".join(lines) + "\n"
            inf.write(test_data)

            result = subprocess.run(exe_path, input=test_data, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding="utf-8", errors="ignore")
            
            print(f"\n第 {i} 組測試答案: ")
            a_line = result.stdout
            print(a_line)

            anf.write(a_line.strip() + "\n")
            inf.write("\n")
            anf.write("\n")

if __name__ == "__main__":
    main()
    input("執行完畢，請按 Enter 鍵退出...")
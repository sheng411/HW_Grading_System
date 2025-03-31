import os
import subprocess

compile_path = r"C:\msys64\mingw64\bin\g++.exe"

def main():
    problem = int(input("題號(預設為0): ").strip()or 0)
    num_tests = int(input("測資數量(預設為10): ")or 10)
    param_count_str = int(input("測試資料行數(預設為1): ")or 1)
    code_path = input("輸入測試程式路徑(預設為當前): ").strip() or os.getcwd()

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

    compile_cmd = [compile_path,"-o", exe_path,cpp_path]
    result = subprocess.run(compile_cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if result.returncode != 0:
        print(f"\n編譯 {os.path.basename(cpp_path)} 時發生錯誤: {result.stderr}")
        return
    else:
        print(f"\n成功編譯 {os.path.basename(cpp_path)} 為 {os.path.basename(exe_path)}.exe")
    
    #print(f"input_path: {input_path}")
    

    #test_inputs = []
    #answers = []
    print("\n請依序輸入測試資料(每輸入一行請按 Enter): ")

    with open(input_path, "w", encoding="utf-8") as inf,open(ans_path, "w", encoding="utf-8") as anf:
        for i in range(1,num_tests+1):
            print(f"\n\n第 {i} 組測試資料: ")
            test_data = ""
            for j in range(param_count_str):
                i_line = input().rstrip()  # 去除換行與多餘空白
                if i_line == "":
                    break
                
                #test_inputs.append(i_line)
                test_data += i_line + "\n"
            inf.write(test_data)

            result = subprocess.run(exe_path,input=test_data,stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True,encoding="utf-8",errors="ignore")    #replace errors="ignore" 
            
            print(f"\n第 {i} 組測試答案: ")
            a_line = result.stdout
            print(a_line)

            anf.write(a_line)
            inf.write("\n")
            anf.write("\n")
            #answers.append(a_line)
            #test_inputs.append("")
            #answers.append("")

    '''
    # 將測試資料寫入 input 檔案中
    with open(input_path, "w", encoding="utf-8") as inf:
        for line in test_inputs:
            inf.write(line + "\n")
    print(f"\n\n測試資料已寫入 {input_filename}")
    
    # 將答案寫入 ans 檔案中
    with open(ans_path, "w", encoding="utf-8") as anf:
        for ans in answers:
            anf.write(ans + "\n")
    print(f"答案已寫入 {ans_filename}\n\n")
    '''
if __name__ == "__main__":
    main()
    input("執行完畢，請按 Enter 鍵退出...")
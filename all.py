import os
import re
import subprocess
import time
import shutil
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

compile_path = r"C:\msys64\mingw64\bin\g++.exe"
test_input_file = "testinput.txt"
total_file_name = "total.txt"
error_student_folder = "0Error"

'''
test input file name: 1input.txt, 2input.txt, ...
answer file name: 1ans.txt, 2ans.txt, ...
testinput.txt: 暫存每次測試的輸入
total.txt: 統計學生的答題狀況
Score_{作業日期}.xlsx: 紀錄該次作業狀況

'''


#find student root
def find_student_root(base_dir):
    subfolders = [d for d in os.listdir(base_dir) if os.path.isdir(os.path.join(base_dir, d))]
    if len(subfolders) == 1:
        #print("subfolders: ",subfolders)
        return os.path.join(base_dir, subfolders[0])
    else:
        max_count = 0
        candidate = None
        for folder in subfolders:
            folder_path = os.path.join(base_dir, folder)
            #print("folder_path: ",folder_path)
            subdirs = [d for d in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, d))]
            #print("subdirs: ",subdirs)
            if len(subdirs) > max_count:
                max_count = len(subdirs)
                candidate = folder_path
        return candidate
#rename student folders
def rename_student_folders_in_root(student_root):
    for folder in os.listdir(student_root):
        folder_path = os.path.join(student_root, folder)
        if os.path.isdir(folder_path):
            # 修改正則表達式以捕獲中文姓名與學號
            m = re.search(r'^([\u4e00-\u9fff]+).*?(\d{8,9}[A-Za-z])(?=_|$)', folder)
            #print("m: ",m)
            if m:
                chinese_name = m.group(1)
                student_id = m.group(2)
                #new_name = f"{chinese_name}_{student_id}"
                #new_name = f"{student_id}_{chinese_name}"
                new_name = f"{student_id}"
                new_path = os.path.join(student_root, new_name)
                if folder != new_name:
                    if os.path.exists(new_path):
                        print(f"新名稱 '{new_name}' 已存在，無法重新命名 '{folder}'")
                    else:
                        print(f"將資料夾 '{folder}' 重新命名為 '{new_name}'")
                        os.rename(folder_path, new_path)
                else:
                    print(f"資料夾 '{folder}' 已符合格式，不做修改。")
            else:
                print(f"資料夾 '{folder}' 中找不到符合格式的中文姓名及學號，跳過重新命名。")
#generate excel
def generate_excel(student_root, check_excel):
    """
    生成一個 Excel 檔案 (StudentList.xlsx) 存放學生資料夾名稱與學號。
    Excel 的 A1 儲存格填入 "dir_name"，B1 填入 "S_id"，
    從 A2 開始，每一列依序填入學生資料夾的新名稱與對應的學號（以底線分隔，第一部分為學號）。
    """
    if check_excel and os.path.exists(excel_file):
        # 開啟既有 Excel
        wb = load_workbook(excel_file)
        ws = wb.active  # 假設處理第一個工作表

        # 收集現有的 S_id（B 欄）到一個 set
        existing_sids = set()
        for row_idx in range(2, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=2).value
            if cell_value:
                existing_sids.add(cell_value)
        
        # 從下一個空白列開始插入新資料
        row = ws.max_row + 1
        print(f"已讀取 {excel_file}，目前已有 {len(existing_sids)} 筆學號紀錄。")
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Student List"
        ws["A1"] = "dir_name"
        ws["B1"] = "S_name"
        ws["C1"] = "Error score"
        ws["D1"] = "Score"
        ws["E1"] = "Note"

        existing_sids = set()
        row = 2
        print(f"建立新的 {excel_file} 檔案")


    for folder in os.listdir(student_root):
        folder_path = os.path.join(student_root, folder)
        if os.path.isdir(folder_path):
            # 假設新名稱格式為 "學號_中文姓名"，底線前為學號
            dir_name = folder
            parts = folder.split("_")
            s_id = parts[0] if parts else ""
            #s_id=s_name

            if s_id==error_student_folder:
                continue
            # 若該學號不在 Excel 紀錄中，則插入新資料
            elif s_id and s_id not in existing_sids:
                ws.cell(row=row, column=1, value=dir_name)  # A 欄：dir_name
                ws.cell(row=row, column=2, value=s_id)      # B 欄：S_id
                row += 1
                existing_sids.add(s_id)

    wb.save(excel_file)
    print(f"Excel 檔案已儲存：{excel_file}")


#move non cpp files
def move_non_cpp_folders(hw_folder_path):
    check_count = 1
    error_folder_path = os.path.join(hw_folder_path, error_student_folder)

    # 如果錯誤資料夾不存在，則建立它
    if not os.path.exists(error_folder_path):
        os.makedirs(error_folder_path)

    # 遍歷 HW_folder 內的學生資料夾
    for student_folder in os.listdir(hw_folder_path):
        student_folder_path = os.path.join(hw_folder_path, student_folder)

        # 跳過錯誤資料夾本身
        if student_folder == error_student_folder:
            continue

        # 確保是資料夾
        if os.path.isdir(student_folder_path):
            has_non_cpp = False  # 用來判斷該學生資料夾內是否有非 .cpp 檔案

            for file in os.listdir(student_folder_path):
                file_path = os.path.join(student_folder_path, file)

                # 檢查是否為檔案，且副檔名不是 .cpp
                if os.path.isfile(file_path) and not file.endswith('.cpp'):
                    has_non_cpp = True
                    break  # 找到非 .cpp 檔案就可提前結束該資料夾的搜尋

            # 如果該學生資料夾內含有非 .cpp 檔案，則移動至錯誤資料夾
            if has_non_cpp:
                target_path = os.path.join(error_folder_path, student_folder)
                
                # 如果目標資料夾已存在（避免衝突），先刪除它
                if os.path.exists(target_path):
                    shutil.rmtree(target_path)

                shutil.move(student_folder_path, target_path)
                print(f"{check_count}. 已將 {student_folder} 移動到 {error_student_folder} 資料夾")
                check_count += 1



#compile and test
def process_student_folder(folder, num_programs,test_num_i):
    for i in range(1, num_programs + 1):
        cpp_filename = f"{i}.cpp"
        cpp_path = os.path.join(folder, cpp_filename)


        #if os.path.basename(folder) == error_student_folder:
        #    print(f"跳過 {cpp_filename}，因為資料夾名稱為 {error_student_folder}")
        #    continue

        if not os.path.isfile(cpp_path):
            print(f"*找不到 {cpp_filename},Pass {i}")
            continue

        # 編譯程式，將執行檔命名為「i」(不含副檔名)
        exe_path = os.path.join(folder, f"{i}")
        #compile_cmd = ["g++", cpp_path, "-o", exe_path]
        compile_cmd = [compile_path,"-o", exe_path,cpp_path]
        #compile_cmd = [compile_path, cpp_path, "-o", exe_path]


        result = subprocess.run(compile_cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if result.returncode != 0:
            print(f"編譯 {os.path.basename(cpp_path)} 時發生錯誤: {result.stderr}")
            continue
        else:
            print(f"成功編譯 {os.path.basename(cpp_path)} 為 {os.path.basename(exe_path)}.exe")

        # 設定測試檔案：
        input_filename = f"{i}input.txt"
        input_file = input_filename
        #test_input_file = "testinput.txt"

        # 輸出結果
        output_filename = f"{i}output.txt"
        output_file = os.path.join(folder, output_filename)

        # clear output_file
        with open(output_file, "w") as outf:
            outf.write("")

        # read input_file
        if not os.path.isfile(input_file):
            print(f"找不到輸入檔 {input_file}，跳過 {cpp_filename}。")
            continue

        with open(input_file, "r") as f:
            #lines = f.readlines()
            lines = [line.rstrip() for line in f if line.strip() != ""]
        f.close()

        
        for line in lines:
            test_data = line.strip()
            if test_data == "":
                continue  # 忽略空行
            '''

        # 測資並寫入
        #test_num_i = 0
        while test_num_i < len(lines):
            try:
                # 讀取測資區塊第一行，代表後續的資料行數
                count = int(lines[test_num_i])
            except ValueError:
                print(f"無法轉換成整數，測資格式錯誤：{lines[test_num_i]}")
                break

            # 檢查是否有足夠的行數：header行 + count 行
            if test_num_i + 1 + count > len(lines):
                print("測資區塊不完整。")
                break
        
            # 取得本次測資區塊所有行（包含第一行）
            test_case_lines = lines[test_num_i : test_num_i + 1 + count]
            # 組合成一個字串，並於最後補上換行符號
            test_data = "\n".join(test_case_lines) + "\n"
            '''
            
            with open(test_input_file, "w") as tif:
                tif.write(test_data)
            tif.close()

            # 使用 testinput.txt 作為標準輸入來執行編譯後的程式
            with open(test_input_file, "r") as tif:
                run_result = subprocess.run(exe_path,stdin=tif,stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True,encoding="utf-8",errors="replace")
            tif.close()
                
            output=run_result.stdout
            #print(repr(output))
            if not output.endswith("\n"):
                output += "\n"
            #print(repr(output))
            
            # 結果寫入
            with open(output_file, "a") as outf:
                outf.write(output)            

            # clear testinput.txt
            with open(test_input_file, "w") as tif:
                tif.write("")
            tif.close()
#比對答案
def comparison_student_data(folder, num_problems, high_num_problems):
    """
    比對資料夾內的 ioutput.txt 與 ians.txt，
    回傳：
      - results: 每題 "O" 或 "X"
      - error_count: 錯誤題數
      - wrong_questions: 所有答錯的題號清單 (list of int)
    """
    results = []
    error_count = 0
    wrong_questions = []  # 用來記錄哪幾題錯誤
    high_start= num_problems - high_num_problems + 1
    high_error_count = 0

    for i in range(1, num_problems + 1):
        output_file = os.path.join(folder, f"{i}output.txt")
        ans_file = os.path.join(os.getcwd(), f"{i}ans.txt")

        if not os.path.exists(output_file):
            print(f"找不到 {i}output.txt")
            results.append("X")
            if i >= high_start:
                high_error_count += 1
            else:
                error_count += 1
            wrong_questions.append(i)
            continue
        if not os.path.exists(ans_file):
            print(f"找不到答案檔 {i}ans.txt，題號 {i}")
            results.append("?")
            if i >= high_start:
                high_error_count += 1
            else:
                error_count += 1
            wrong_questions.append(i)
            continue

        with open(output_file, "r", encoding="utf-8", errors="replace") as f_out:
            student_output = f_out.read().strip()
        f_out.close()

        with open(ans_file, "r", encoding="utf-8") as f_ans:
            correct_output = f_ans.read().strip()
        f_ans.close()

        if student_output == correct_output:
            results.append("O")
        elif i >= high_start:
            print(f"第 {i} 題與答案不符")
            results.append("X")
            high_error_count += 1
            wrong_questions.append(i)
        else:
            print(f"第 {i} 題與答案不符")
            results.append("X")
            error_count += 1
            wrong_questions.append(i)

    return results, error_count, high_error_count, wrong_questions
#寫入E欄(錯誤題目)
def write_errors_to_excel(student, wrong_questions):
    """
    若 wrong_questions 不為空，則在 Excel (StudentList.xlsx) 裡，
    尋找 A 欄 (dir_name) 與 student 相同的列，將
    '第 1,2,3 題錯誤' 寫入 E 欄 (第 5 欄)。
    """
    if not wrong_questions:
        # 全對就不寫任何東西
        return

    if not os.path.exists(excel_file):
        print(f"Excel 檔 {excel_file} 不存在，無法更新。")
        return

    wb = load_workbook(excel_file)
    ws = wb.active  # 假設要更新第一個工作表

    # 尋找 A 欄符合 student 的列
    found_row = None
    for row_idx in range(2, ws.max_row + 1):
        dir_name_cell = ws.cell(row=row_idx, column=1).value
        if dir_name_cell == student:
            found_row = row_idx
            break

    if found_row:
        # 組合字串：例如 "第 1,2,3 題錯誤"
        error_str = "第 " + ",".join(str(q) for q in wrong_questions) + " 題錯誤"
        ws.cell(found_row, 5).value = error_str  # 第 5 欄 (E 欄)
        print(f"已在 Excel 中將 {student} 的錯題寫入 E 欄：{error_str}")
    else:
        print(f"未在 Excel 中找到 {student}，無法寫入錯題資訊。")

    wb.save(excel_file)
#寫入C、D欄(錯誤題數、分數)
def add_excel(student, total_problems, high_num_problems, error_count,score_ballast ,high_error_count ,high_score_ballast):
    """
    參數:學生/總題數/高分題數/基本錯題/基本配分/高分錯題/高分配分
    開啟已存在的 StudentList.xlsx，尋找 A 欄與 student 相同的列，
    若找到則將 error_count 寫入 C 欄，(total_problems - error_count) 寫入 D 欄。
    """
    base_num_problems = total_problems - high_num_problems

    if not os.path.exists(excel_file):
        print(f"Excel 檔 {excel_file} 不存在，無法更新。")
        return

    # 讀取既有的 Excel
    wb = load_workbook(excel_file)
    ws = wb.active  # 假設要更新第一個工作表
    
    # 從第 2 列開始尋找（跳過標題列）
    found_row = None
    for row in range(2, ws.max_row + 1):
        dir_name_cell = ws.cell(row=row, column=1).value  # A 欄
        if dir_name_cell == student:
            found_row = row
            break

    if found_row:
        # C 欄 (column=3) 寫入錯誤題數
        ws.cell(found_row, 3).value = (error_count+high_error_count)
        final_score = ((base_num_problems - error_count)*score_ballast)+((high_num_problems - high_error_count)*high_score_ballast)
        print(f"Score: (({base_num_problems} - {error_count})*{score_ballast})+(({high_num_problems} - {high_error_count})*{high_score_ballast})={final_score}")
        # D 欄 (column=4) 寫入總題數 - 錯題數
        ws.cell(found_row, 4).value = final_score
        print(f"\n已更新 {student} 的錯誤題數 = {error_count}，分數 = {final_score}")
    else:
        print(f"未在 Excel 中找到 {student}，無法更新。")

    wb.save(excel_file)


#setting excel format
def format_excel(excel_file):
    """
    開啟指定的 Excel 檔案，將所有儲存格的字體設定為 size=12，
    並自動調整每一欄的寬度，使內容能夠完整顯示。
    """
    if not os.path.exists(excel_file):
        print(f"找不到檔案：{excel_file}")
        return

    # 讀取既有 Excel
    wb = load_workbook(excel_file)
    ws = wb.active  # 假設要處理第一個工作表

    # 1) 設定字體大小為 12
    for row in ws.iter_rows():
        for cell in row:
            # 如果想強制所有儲存格都設定字體大小，即使它是空值也可
            # 若只想設定有值的儲存格，可以檢查 cell.value
            cell.font = Font(size=12)

    # 2) 模擬「自動欄寬」
    # 逐欄找出該欄最長字串長度，然後設定 column_dimensions
    for col_idx, col_cells in enumerate(ws.iter_cols(min_col=1,
                                                     max_col=ws.max_column,
                                                     min_row=1,
                                                     max_row=ws.max_row), start=1):
        max_length = 0
        for cell in col_cells:
            # 若該儲存格有值，計算其長度
            if cell.value is not None:
                length = len(str(cell.value))
                if length > max_length:
                    max_length = length
        # 取出該欄的字母代號 (A, B, C...)
        col_letter = get_column_letter(col_idx)
        # 設定欄寬：可再自行加減，避免過於擁擠
        ws.column_dimensions[col_letter].width = max_length + 5

    # 存檔
    wb.save(excel_file)
    print(f"已完成對 {excel_file} 的字體與欄寬調整。")


def main():
    global excel_file
    # 詢問使用者本次要執行幾個程式
    try:
        num_problems = int(input("檢測程式數量? "))
    except ValueError:
        print("請輸入有效的整數。")
        return
    score_ballast = int(input("基本題配分: "))
    high_num_problems = int(input("困難題目數量: "))
    high_score_ballast = int(input("困難題目配分: "))
    test_num_i = int(input("每次測試資料行數(預設為1): "))
    

    selection=input("作業編號(eg.02261): ")
    excel_file = f"Score_{selection}.xlsx"
    
    print("\n\n--------------- START INSPECION ---------------")
    start = time.perf_counter()

    #rename
    base_dir = os.getcwd()      #當前目錄
    student_root = find_student_root(base_dir)  #學生資料夾目錄
    #print("base_dir: ",base_dir)
    #print("student_root: ",student_root)
    
    rename_student_folders_in_root(student_root)

    if os.path.exists(excel_file):
        print(f"{excel_file} is True")
        check_excel = 1
    else:
        print(f"{excel_file} is False")
        check_excel = 0
    
    # 生成 Excel 學生清單
    generate_excel(student_root,check_excel)

    print("\n\n--------------- CHECK .cpp FILE ---------------")

    #check non cpp files
    move_non_cpp_folders(student_root)
    print("\n\n--------------- TESTING ---------------")

    num=0
    items = [os.path.join(student_root, d) for d in os.listdir(student_root) if os.path.isdir(os.path.join(student_root, d))]
    total_file_path = os.path.join(base_dir, total_file_name)
    with open(total_file_path, "w", encoding="utf-8") as total_file:
        for item in items:
            #print("item: ",item)
            if os.path.isdir(item):
                if os.path.basename(item) == error_student_folder:
                    #print(f"跳過 {item}，因為資料夾名稱為 {error_student_folder}")
                    continue
                num+=1
                print(f"\n\n{num}.處理資料夾：{os.path.basename(item)}")
                process_student_folder(item, num_problems,test_num_i)

                student = os.path.basename(item)
                #print(f"{num}. 學生 {student}")
                results, error_count, high_error_count, wrong_questions = comparison_student_data(item, num_problems, high_num_problems)
                #print("wrong_questions: ",wrong_questions)
                write_errors_to_excel(student, wrong_questions)
                result_str = " ".join(results)
                line = f"{student}----- {result_str} -----基本題錯 {error_count} 題 進階題錯 {high_error_count} 題\n"
                #學生/總題數/高分題數/基本錯題/基本配分/高分錯題/高分配分
                add_excel(student,num_problems, high_num_problems,error_count,score_ballast, high_error_count,high_score_ballast)

                total_file.write(line)
                print(f"結果: {line.strip()}\n\n")
    total_file.close()

            
    print("\n\n--------------- END INSPECTION ---------------\n")
    
 
    #設定Excel格式
    format_excel(excel_file)

    end = time.perf_counter()
    print(f"\n\n執行時間: {end - start:.2f} 秒")
    print(f"平均每位同學處裡時間: {(end - start)/num:.2f} 秒\n\n")

if __name__ == "__main__":
    main()
    input("執行完畢，請按 Enter 鍵退出...")
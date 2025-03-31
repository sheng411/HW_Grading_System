from fpdf import FPDF
import os
import shutil
import warnings

warnings.filterwarnings("ignore")
class PDF(FPDF):
    def __init__(self, student_id):
        super().__init__()  # 初始化 FPDF
        self.student_id = student_id  # 存入物件屬性

    def header(self):
        self.set_font('DFKai-SB', '', 12)  # 確保字型已載入
        title = self.student_id
        title_w = self.get_string_width(title) + 6
        doc_w = self.w
        self.set_x((doc_w - title_w) / 2)
        self.cell(title_w, 10, title, border=1, ln=1, align='C')
        self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font('DFKai-SB', '', 8)
        self.cell(0, 10, f'{self.page_no()}', align='C')


def txt2pdf(student_id, txt_file,item,hw_dir_path):
    pdf_file = f"{student_id}.pdf"
    pdf_output_path=os.path.join(item,pdf_file)
    hw_pdf_dir=os.path.join(hw_dir_path,"0All_pdf")
    if not os.path.exists(hw_pdf_dir):
        os.makedirs(hw_pdf_dir)  # 如果資料夾不存在，則建立它
        print(f"已建立資料夾: {hw_pdf_dir}")
    else:
        print(f"資料夾 {hw_pdf_dir} 已存在，將覆蓋檔案。")

    # 檢查字型檔案是否存在
    font_path = "kaiu.ttf"
    if not os.path.exists(font_path):
        raise FileNotFoundError(f"字型檔案 '{font_path}' 找不到，請確認檔案已放在相同目錄下。")

    # 建立 PDF 物件
    pdf = PDF(student_id)

    # 載入標楷體字型
    pdf.add_font('DFKai-SB', '', font_path, uni=True)  # 這一定要在 `set_font()` 之前

    # 設定字型
    pdf.set_font("DFKai-SB", size=12)

    pdf.add_page()

    # 讀取純文字檔並寫入 PDF
    with open(txt_file, "r", encoding="utf-8") as file:
        for line in file:
            pdf.multi_cell(0, 5, line.strip())

    pdf.output(pdf_output_path)
    shutil.copy(pdf_output_path, hw_pdf_dir)


    print("PDF OK!")


if __name__ == "__main__":
    student_id = "110011"
    txt_file = "total.txt"
    txt2pdf(student_id, txt_file)
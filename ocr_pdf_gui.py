import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pytesseract
from pdf2image import convert_from_path
from fpdf import FPDF
from PIL import Image
import tempfile
import threading

# 必要に応じてTesseractのパスを明示（Windows向け、要調整）
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def ocr_pdf(input_pdf, output_pdf, lang='jpn+eng', progress_var=None):
    images = convert_from_path(input_pdf)
    pdf = FPDF(unit='pt')
    total_pages = len(images)
    for i, image in enumerate(images):
        if progress_var is not None:
            progress_var.set(int((i+1)/total_pages*100))
        ocr_result = pytesseract.image_to_string(image, lang=lang)
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as fp:
            temp_img = fp.name
            image.save(temp_img)
        w, h = image.size
        pdf.add_page(orientation='P' if h >= w else 'L')
        pdf.image(temp_img, 0, 0, w, h)
        pdf.set_xy(10, h-80)
        pdf.set_font("Arial", size=10)
        pdf.multi_cell(0, 10, ocr_result)
        os.remove(temp_img)
    pdf.output(output_pdf)

def run_ocr(input_pdf, output_pdf, lang, progress_var, btn):
    try:
        btn.config(state="disabled")
        ocr_pdf(input_pdf, output_pdf, lang=lang, progress_var=progress_var)
        messagebox.showinfo("完了", f"OCR PDF作成が完了しました。\n{output_pdf}")
    except Exception as e:
        messagebox.showerror("エラー", f"OCR PDF作成中にエラー: {e}")
    finally:
        btn.config(state="normal")
        progress_var.set(0)

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF OCRツール")
        self.geometry("450x250")

        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.lang = tk.StringVar(value='jpn+eng')
        self.progress = tk.IntVar()

        # 入力ラベル＆テキスト
        tk.Label(self, text="入力PDFファイル:").pack(anchor=tk.W, pady=(10,0), padx=10)
        frame1 = tk.Frame(self)
        frame1.pack(fill=tk.X, padx=10)
        tk.Entry(frame1, textvariable=self.input_path, width=35).pack(side=tk.LEFT, expand=True, fill=tk.X)
        tk.Button(frame1, text="参照", command=self.select_input).pack(side=tk.LEFT, padx=5)

        # 出力ラベル＆テキスト
        tk.Label(self, text="出力PDFファイル:").pack(anchor=tk.W, pady=(10,0), padx=10)
        frame2 = tk.Frame(self)
        frame2.pack(fill=tk.X, padx=10)
        tk.Entry(frame2, textvariable=self.output_path, width=35).pack(side=tk.LEFT, expand=True, fill=tk.X)
        tk.Button(frame2, text="参照", command=self.select_output).pack(side=tk.LEFT, padx=5)

        # 言語選択
        tk.Label(self, text="OCR言語 (例:jpn, eng, jpn+eng):").pack(anchor=tk.W, pady=(10,0), padx=10)
        tk.Entry(self, textvariable=self.lang, width=10).pack(anchor=tk.W, padx=10)

        # 進捗バー
        self.progressbar = tk.Scale(self, from_=0, to=100, orient=tk.HORIZONTAL, variable=self.progress, length=350)
        self.progressbar.pack(pady=(10,0))

        # 実行ボタン
        self.btn = tk.Button(self, text="OCRしてPDF出力", command=self.start)
        self.btn.pack(pady=15)

    def select_input(self):
        filename = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if filename:
            self.input_path.set(filename)
            # 自動で出力名も変更
            base, ext = os.path.splitext(filename)
            self.output_path.set(base + "_ocr.pdf")

    def select_output(self):
        filename = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if filename:
            self.output_path.set(filename)

    def start(self):
        if not self.input_path.get() or not self.output_path.get():
            messagebox.showwarning("ファイル未指定", "入力PDF・出力PDFファイルを指定してください。")
            return
        t = threading.Thread(
            target=run_ocr,
            args=(self.input_path.get(), self.output_path.get(), self.lang.get(), self.progress, self.btn)
        )
        t.start()

if __name__ == "__main__":
    app = App()
    app.mainloop()

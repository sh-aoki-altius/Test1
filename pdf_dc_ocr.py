import os
import tkinter as tk
from tkinter import (
    Tk, Frame, Button, Label, Listbox, SINGLE, END,
    filedialog, messagebox, simpledialog, Checkbutton, IntVar
)
from PyPDF2 import PdfReader, PdfWriter
import tempfile
import shutil

# --- OCR系ライブラリ ---
from pdf2image import convert_from_path
from PIL import Image
import pytesseract
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

# --- Popplerパス（必要に応じて書き換えてください。Windowsのみ）---
POPPLER_PATH = None  # 例: r"C:\tools\poppler-23.11.0\Library\bin"


def pdf_page_to_searchable_pdf_page(pdf_path, page_index, lang='jpn+eng'):
    """
    1ページ単位で，画像からOCRをかけサーチャブルPDF1ページを一時ファイルとして返す
    """
    images = convert_from_path(
        pdf_path,
        first_page=page_index+1,
        last_page=page_index+1,
        poppler_path=POPPLER_PATH,
        dpi=300
    )
    image = images[0]
    ocr_text = pytesseract.image_to_string(image, lang=lang)
    # 一時PDF
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_pdf:
        width, height = image.size
        c = canvas.Canvas(tmp_pdf.name, pagesize=(width, height))
        # 背景画像（PDF元の見た目）
        c.drawImage(ImageReader(image), 0, 0, width=width, height=height)
        # 透明テキスト（OCR結果）
        c.setFont("Helvetica", 10)
        c.setFillColorRGB(0, 0, 0, alpha=0)  # 透明フォント（不可視でサーチャブルになる）
        # OCRテキストを下部に一括埋め込み
        lines = ocr_text.split('\n')
        y = 30
        for line in lines:
            c.drawString(10, y, line)
            y += 12
            if y > height - 20: break
        c.showPage()
        c.save()
    return tmp_pdf.name

class PDFToolGUI:
    def __init__(self, root):
        self.root = root
        self.ocr_var_merge = IntVar()
        self.ocr_var_split = IntVar()
        self.root.title("PDF結合/分割＋OCRツール")
        self.root.geometry("520x480")
        self.root.configure(bg="#fff")

        self.password_cache = {}

        self.mode_frame = Frame(root, bg="#fff")
        self.mode_frame.pack(pady=5)

        self.merge_btn = Button(self.mode_frame, text="🔗 PDF結合", font=("Meiryo", 10, "bold"), width=20, command=self.show_merge_mode)
        self.split_btn = Button(self.mode_frame, text="✂ PDF分割", font=("Meiryo", 10, "bold"), width=20, command=self.show_split_mode)
        self.merge_btn.grid(row=0, column=0, padx=5)
        self.split_btn.grid(row=0, column=1, padx=5)

        self.merge_frame = Frame(root, bg="#fff")
        self.split_frame = Frame(root, bg="#fff")

        self.setup_merge_frame()
        self.setup_split_frame()

        self.pdf_paths = []
        self.drag_data = {"widget": None, "index": None}
        self.current_mode = None

        self.ocr_var_merge = IntVar()
        self.ocr_var_split = IntVar()

        self.show_merge_mode()

    def update_mode_buttons(self, active):
        if active == "merge":
            self.merge_btn.config(bg="#2196F3", fg="white", relief="sunken")
            self.split_btn.config(bg="SystemButtonFace", fg="black", relief="raised")
        elif active == "split":
            self.split_btn.config(bg="#FF9800", fg="white", relief="sunken")
            self.merge_btn.config(bg="SystemButtonFace", fg="black", relief="raised")

    def show_merge_mode(self):
        if self.current_mode == "merge": return
        self.split_frame.pack_forget()
        self.merge_frame.pack(fill="both", expand=True)
        self.update_mode_buttons("merge")
        self.current_mode = "merge"

    def show_split_mode(self):
        if self.current_mode == "split": return
        self.merge_frame.pack_forget()
        self.split_frame.pack(fill="both", expand=True)
        self.update_mode_buttons("split")
        self.current_mode = "split"

    def setup_merge_frame(self):
        Label(self.merge_frame, text="① PDFが入ったフォルダを選択", font=("Meiryo", 10, "bold"), bg="#fff").pack(pady=10)
        Button(self.merge_frame, text="📂 フォルダ選択", font=("Meiryo", 8, "bold"), command=self.select_folder).pack()

        self.listbox = Listbox(self.merge_frame, selectmode=SINGLE, width=60)
        self.listbox.pack(pady=10)
        self.listbox.bind("<Enter>", lambda e: self.listbox.config(cursor="hand2"))
        self.listbox.bind("<Leave>", lambda e: self.listbox.config(cursor=""))
        self.listbox.bind("<ButtonPress-1>", self.on_drag_start)
        self.listbox.bind("<B1-Motion>", self.on_drag_motion)
        self.listbox.bind("<ButtonRelease-1>", self.on_drag_drop)

        Label(self.merge_frame, text="② 並び順をドラッグで調整 → 結合", font=("Meiryo", 10, "bold"), bg="#fff").pack(pady=10)

        c1 = Checkbutton(self.merge_frame, text="OCR(文字認識)してサーチャブルPDF化", variable=self.ocr_var_merge, bg="#fff")
        c1.pack()
        Button(self.merge_frame, text="▶ この順で結合", font=("Meiryo", 8, "bold"), command=self.merge_pdfs, bg="#4CAF50", fg="white").pack(pady=5)

    def select_folder(self):
        folder = filedialog.askdirectory(title="PDFが入ったフォルダを選択")
        if not folder: return
        files = [f for f in os.listdir(folder) if f.lower().endswith(".pdf")]
        files.sort()
        self.pdf_paths = [os.path.join(folder, f) for f in files]
        self.listbox.delete(0, END)
        for f in files:
            self.listbox.insert(END, f)

    def on_drag_start(self, event):
        index = self.listbox.nearest(event.y)
        if index >= 0:
            self.drag_data = {"widget": self.listbox, "index": index}

    def on_drag_motion(self, event):
        pass

    def on_drag_drop(self, event):
        widget = self.drag_data["widget"]
        from_index = self.drag_data["index"]
        to_index = widget.nearest(event.y)
        if from_index is None or to_index is None or from_index == to_index: return
        text = widget.get(from_index)
        widget.delete(from_index)
        widget.insert(to_index, text)
        widget.select_set(to_index)
        self.pdf_paths[from_index], self.pdf_paths[to_index] = self.pdf_paths[to_index], self.pdf_paths[from_index]
        self.drag_data = {"widget": None, "index": None}

    def get_pdf_reader(self, path):
        reader = PdfReader(path)
        if reader.is_encrypted:
            cached = self.password_cache.get(path)
            if cached and reader.decrypt(cached) != 0:
                return reader
            while True:
                pwd = simpledialog.askstring("パスワード要求", f"{os.path.basename(path)} のパスワードを入力：", show="*")
                if not pwd:
                    raise Exception(f"{os.path.basename(path)} の読み込みを中止しました")
                if reader.decrypt(pwd) != 0:
                    self.password_cache[path] = pwd
                    return reader
                else:
                    messagebox.showerror("パスワードエラー", f"パスワードが正しくありません: {os.path.basename(path)}")
        return reader

    def merge_pdfs(self):
        if not self.pdf_paths:
            messagebox.showwarning("警告", "まずPDFを含むフォルダを選択してください")
            return

        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], title="結合後のPDFファイルを保存")
        if not save_path: return

        ocr = self.ocr_var_merge.get() == 1

        try:
            writer = PdfWriter()
            tempfiles = []
            for path in self.pdf_paths:
                reader = self.get_pdf_reader(path)
                for i, page in enumerate(reader.pages):
                    if ocr:
                        # 各ページOCR化（サーチャブルPDFページを一時生成してPyPDF2で読込）
                        temp_pdf = pdf_page_to_searchable_pdf_page(path, i)
                        tempfiles.append(temp_pdf)
                        ocr_reader = PdfReader(temp_pdf)
                        writer.add_page(ocr_reader.pages[0])
                    else:
                        writer.add_page(page)
            with open(save_path, "wb") as f:
                writer.write(f)
            for tf in tempfiles:
                try:
                    os.unlink(tf)
                except:
                    pass
            if ocr:
                messagebox.showinfo("完了", f"{len(self.pdf_paths)}ファイルをOCRサーチャブル結合しました。")
            else:
                messagebox.showinfo("完了", f"{len(self.pdf_paths)}ファイルを結合しました。")
        except Exception as e:
            messagebox.showerror("エラー", str(e))

    def setup_split_frame(self):
        Label(self.split_frame, text="分割したいページ範囲（例：1,3,5-7）", font=("Meiryo", 8), bg="#fff").pack(pady=5)
        self.page_range_entry = tk.Entry(self.split_frame, width=30)
        self.page_range_entry.pack(pady=5)

        Label(self.split_frame, text="分割したいPDFを選択", font=("Meiryo", 10, "bold"), bg="#fff").pack(pady=10)
        Button(self.split_frame, text="📄 PDF選択（複数可）", font=("Meiryo", 8, "bold"), command=self.split_pdfs).pack(pady=5)
        c2 = Checkbutton(self.split_frame, text="分割時ページごとにOCR(サーチャブルPDF化)", variable=self.ocr_var_split, bg="#fff")
        c2.pack()

    def parse_page_ranges(self, page_range_str, total_pages):
        pages = set()
        parts = page_range_str.split(",")
        for part in parts:
            part = part.strip()
            if "-" in part:
                try:
                    start, end = map(int, part.split("-"))
                    pages.update(range(max(1, start), min(total_pages, end)+1))
                except:
                    continue
            else:
                try:
                    p = int(part)
                    if 1 <= p <= total_pages:
                        pages.add(p)
                except:
                    continue
        return sorted(p - 1 for p in pages)  # 0-indexed

    def split_pdfs(self):
        input_files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")], title="分割したいPDFファイルを選択")
        if not input_files: return
        output_dir = filedialog.askdirectory(title="出力先フォルダを選択してください")
        if not output_dir: return
        page_range_str = self.page_range_entry.get().strip()
        ocr = self.ocr_var_split.get() == 1

        try:
            total_split_files = 0
            tempfiles = []
            for path in input_files:
                base_name = os.path.splitext(os.path.basename(path))[0]
                reader = self.get_pdf_reader(path)
                total_pages = len(reader.pages)
                pages = self.parse_page_ranges(page_range_str, total_pages) if page_range_str else list(range(total_pages))
                for i in pages:
                    writer = PdfWriter()
                    if ocr:
                        temp_pdf = pdf_page_to_searchable_pdf_page(path, i)
                        tempfiles.append(temp_pdf)
                        ocr_reader = PdfReader(temp_pdf)
                        writer.add_page(ocr_reader.pages[0])
                    else:
                        writer.add_page(reader.pages[i])
                    out_path = os.path.join(output_dir, f"{base_name}_{i+1}.pdf")
                    with open(out_path, "wb") as f:
                        writer.write(f)
                total_split_files += len(pages)
            for tf in tempfiles:
                try:
                    os.unlink(tf)
                except:
                    pass
            if ocr:
                messagebox.showinfo("完了", f"{len(input_files)}ファイル、合計{total_split_files}ページをOCRサーチャブルPDFとして分割しました。")
            else:
                messagebox.showinfo("完了", f"{len(input_files)}ファイル、合計{total_split_files}ページに分割しました。")
        except Exception as e:
            messagebox.showerror("エラー", str(e))


if __name__ == "__main__":
    root = Tk()
    app = PDFToolGUI(root)
    root.mainloop()

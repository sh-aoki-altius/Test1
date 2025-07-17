import os
import tkinter as tk
from tkinter import (
    Tk, Frame, Button, Label, Listbox, SINGLE, END,
    filedialog, messagebox, simpledialog
)
from PyPDF2 import PdfReader, PdfWriter

class PDFToolGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF結合/分割ツール")
        self.root.geometry("500x420")
        self.root.configure(bg="#fff")

        self.password_cache = {}  # パスワード記憶用

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
            if cached and reader.decrypt(cached):
                return reader
            while True:
                pwd = simpledialog.askstring("パスワード要求", f"{os.path.basename(path)} のパスワードを入力：", show="*")
                if not pwd:
                    raise Exception(f"{os.path.basename(path)} の読み込みを中止しました")
                if reader.decrypt(pwd):
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

        try:
            writer = PdfWriter()
            for path in self.pdf_paths:
                reader = self.get_pdf_reader(path)
                for page in reader.pages:
                    writer.add_page(page)
            with open(save_path, "wb") as f:
                writer.write(f)
            messagebox.showinfo("完了", f"{len(self.pdf_paths)}ファイルを結合しました。")
        except Exception as e:
            messagebox.showerror("エラー", str(e))

    def setup_split_frame(self):
        Label(self.split_frame, text="分割したいページ範囲（例：1,3,5-7）", font=("Meiryo", 8), bg="#fff").pack(pady=5)
        self.page_range_entry = tk.Entry(self.split_frame, width=30)
        self.page_range_entry.pack(pady=5)

        Label(self.split_frame, text="分割したいPDFを選択", font=("Meiryo", 10, "bold"), bg="#fff").pack(pady=10)
        Button(self.split_frame, text="📄 PDF選択（複数可）", font=("Meiryo", 8, "bold"), command=self.split_pdfs).pack(pady=5)

    def parse_page_ranges(self, page_range_str, total_pages):
        pages = set()
        parts = page_range_str.split(",")
        for part in parts:
            part = part.strip()
            if "-" in part:
                try:
                    start, end = map(int, part.split("-"))
                    pages.update(range(max(1, start), min(total_pages, end)+1))
                except: continue
            else:
                try:
                    p = int(part)
                    if 1 <= p <= total_pages:
                        pages.add(p)
                except: continue
        return sorted(p - 1 for p in pages)  # 0-indexed

    def split_pdfs(self):
        input_files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")], title="分割したいPDFファイルを選択")
        if not input_files: return
        output_dir = filedialog.askdirectory(title="出力先フォルダを選択してください")
        if not output_dir: return
        page_range_str = self.page_range_entry.get().strip()

        try:
            total_split_files = 0
            for path in input_files:
                base_name = os.path.splitext(os.path.basename(path))[0]
                reader = self.get_pdf_reader(path)
                total_pages = len(reader.pages)
                pages = self.parse_page_ranges(page_range_str, total_pages) if page_range_str else list(range(total_pages))
                for i in pages:
                    writer = PdfWriter()
                    writer.add_page(reader.pages[i])
                    out_path = os.path.join(output_dir, f"{base_name}_{i+1}.pdf")
                    with open(out_path, "wb") as f:
                        writer.write(f)
                total_split_files += len(pages)
            messagebox.showinfo("完了", f"{len(input_files)}ファイル、合計{total_split_files}ページを分割しました。")
        except Exception as e:
            messagebox.showerror("エラー", str(e))

if __name__ == "__main__":
    root = Tk()
    app = PDFToolGUI(root)
    root.mainloop()

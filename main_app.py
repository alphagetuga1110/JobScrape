# 1_main_app.py

import threading
import queue
import tkinter as tk
from tkinter import ttk, filedialog, Toplevel, Listbox, Button, Label, StringVar, END, scrolledtext
import openpyxl

# 分割したファイルをインポート
from excel_writer import append_to_excel
from scrapers import doda_scraper, workport_scraper


def choose_sheet_gui(parent, sheet_names):
    # (この関数の中身は変更ありません)
    result = {"sheet": None}
    window = Toplevel(parent)
    window.title("シートを選択")
    window.geometry("300x400")
    window.attributes("-topmost", True)
    Label(window, text="追記したいシートを選んで決定してください:", pady=10).pack()
    listbox = Listbox(window, height=15)
    for name in sheet_names:
        listbox.insert(END, name)
    listbox.pack(pady=5, padx=10, fill="both", expand=True)
    def on_select():
        selected_indices = listbox.curselection()
        if selected_indices:
            result["sheet"] = listbox.get(selected_indices[0])
            window.destroy()
    Button(window, text="このシートに決定", command=on_select).pack(pady=10)
    parent.wait_window(window)
    return result["sheet"]

class ScrapingApp:
    def __init__(self, master):
        self.master = master
        master.title("統合スクレイピングツール")
        master.geometry("600x550")

        self.url_var = StringVar()
        self.count_var = StringVar(value="10")
        self.filepath_var = StringVar()
        self.sheet_var = StringVar()
        self.site_var = StringVar(value="DODA")
        
        # サイト名と、それに対応するスクレイピング関数を辞書で管理
        self.scraper_functions = {
            "DODA": doda_scraper.scrape,
            "Workport": workport_scraper.scrape
        }

        self.create_widgets()
        self.msg_queue = queue.Queue()
        self.process_queue()

    def create_widgets(self):
        # (GUIの見た目を作る部分は変更ありません)
        site_frame = ttk.Frame(self.master, padding="10")
        site_frame.pack(fill="x")
        ttk.Label(site_frame, text="対象サイト:").pack(side="left", padx=5)
        doda_rb = ttk.Radiobutton(site_frame, text="DODA", variable=self.site_var, value="DODA")
        doda_rb.pack(side="left", padx=5)
        workport_rb = ttk.Radiobutton(site_frame, text="Workport", variable=self.site_var, value="Workport")
        workport_rb.pack(side="left", padx=5)
        input_frame = ttk.Frame(self.master, padding="10")
        input_frame.pack(fill="x")
        ttk.Label(input_frame, text="一覧URL:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.url_entry = ttk.Entry(input_frame, textvariable=self.url_var, width=60)
        self.url_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5, columnspan=2)
        ttk.Label(input_frame, text="取得件数:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.count_entry = ttk.Entry(input_frame, textvariable=self.count_var, width=10)
        self.count_entry.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        ttk.Label(input_frame, text="Excelファイル:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.file_entry = ttk.Entry(input_frame, textvariable=self.filepath_var, width=60, state="readonly")
        self.file_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        self.file_button = ttk.Button(input_frame, text="ファイル選択...", command=self.select_file)
        self.file_button.grid(row=2, column=2, padx=5, pady=5)
        ttk.Label(input_frame, text="選択シート:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.sheet_label = ttk.Label(input_frame, textvariable=self.sheet_var, font=("Yu Gothic UI", 9, "bold"))
        self.sheet_label.grid(row=3, column=1, sticky="w", padx=5, pady=5)
        input_frame.columnconfigure(1, weight=1)
        self.start_button = ttk.Button(self.master, text="スクレイピング開始", command=self.start_scraping_thread)
        self.start_button.pack(pady=10)
        log_frame = ttk.Frame(self.master, padding="10")
        log_frame.pack(fill="both", expand=True)
        ttk.Label(log_frame, text="実行ログ:").pack(anchor="w")
        self.log_area = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=15)
        self.log_area.pack(fill="both", expand=True)
        self.log_area.configure(state='disabled')

    def select_file(self):
        # (この関数の中身は変更ありません)
        filepath = filedialog.askopenfilename(title="Excelファイルを選択", filetypes=[("Excel ファイル", "*.xlsx")])
        if filepath:
            self.filepath_var.set(filepath)
            try:
                workbook = openpyxl.load_workbook(filepath, read_only=True)
                sheet_names = workbook.sheetnames
                target_sheet = choose_sheet_gui(self.master, sheet_names)
                if target_sheet:
                    self.sheet_var.set(target_sheet)
                else:
                    self.sheet_var.set("（シートが選択されませんでした）")
            except Exception as e:
                self.log_message(f"Excelファイル読み込みエラー: {e}")

    def start_scraping_thread(self):
        # (この関数の中身は変更ありません)
        url = self.url_var.get().strip()
        count_str = self.count_var.get().strip()
        filepath = self.filepath_var.get()
        sheet = self.sheet_var.get()
        if not (url and count_str and filepath and sheet and "（" not in sheet):
            self.log_message("エラー: 全ての項目を入力・選択してください。")
            return
        try:
            count = int(count_str)
            if not (1 <= count <= 50): raise ValueError
        except ValueError:
            self.log_message("エラー: 取得件数は1～50の半角数字で入力してください。")
            return
        self.start_button.config(state="disabled")
        self.log_area.configure(state='normal')
        self.log_area.delete('1.0', tk.END)
        self.log_area.configure(state='disabled')
        self.log_message("--- 処理を開始します ---")
        thread = threading.Thread(target=self.scraping_worker, args=(url, count, filepath, sheet), daemon=True)
        thread.start()

    def scraping_worker(self, url, count, filepath, sheet):
        site = self.site_var.get()
        self.msg_queue.put(f"サイト「{site}」のスクレイピングを開始します。")

        # 辞書を使って、選択されたサイトに対応する関数を呼び出す
        scraper_function = self.scraper_functions[site]
        all_data = scraper_function(url, count, self.msg_queue)
        
        if all_data:
            unique_data = []
            seen_companies = set()
            for row in all_data:
                company_name = row[0]
                if company_name not in seen_companies:
                    unique_data.append(row)
                    seen_companies.add(company_name)
            self.msg_queue.put(f"\n取得した{len(all_data)}件から、重複を除いた{len(unique_data)}件をExcelに追記します。")
            
            # 分割したappend_to_excel関数を呼び出す
            append_to_excel(filepath, sheet, unique_data, self.msg_queue)
        else:
            self.msg_queue.put("スクレイピングで取得できたデータがありませんでした。")
        self.msg_queue.put("--- 全ての処理が完了しました ---")
        self.msg_queue.put("ENABLE_BUTTON")

    def process_queue(self):
        # (この関数の中身は変更ありません)
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                if msg == "ENABLE_BUTTON":
                    self.start_button.config(state="normal")
                else:
                    self.log_message(msg)
        except queue.Empty:
            pass
        self.master.after(100, self.process_queue)

    def log_message(self, msg):
        # (この関数の中身は変更ありません)
        self.log_area.configure(state='normal')
        self.log_area.insert(tk.END, msg + "\n")
        self.log_area.see(tk.END)
        self.log_area.configure(state='disabled')

if __name__ == "__main__":
    root = tk.Tk()
    app = ScrapingApp(root)
    root.mainloop()
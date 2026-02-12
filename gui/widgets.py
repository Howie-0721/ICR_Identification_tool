"""
自定義 GUI 元件
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
import os


class ConfigForm(ttk.LabelFrame):
    """設定表單元件"""

    def __init__(self, parent, title, fields, **kwargs):
        super().__init__(parent, text=title, **kwargs)
        self.fields = fields
        self.variables = {}
        self._create_form()

    def _create_form(self):
        """創建表單"""
        for i, (label_text, var_name, field_type, default_value) in enumerate(self.fields):
            # 標籤
            ttk.Label(self, text=f"{label_text}:").grid(row=i, column=0, sticky=tk.W, padx=5, pady=5)

            # 變數
            if field_type == "password":
                var = tk.StringVar(value=default_value)
                entry = ttk.Entry(self, textvariable=var, show="*")
            else:
                var = tk.StringVar(value=default_value)
                entry = ttk.Entry(self, textvariable=var)

            entry.grid(row=i, column=1, sticky=tk.EW, padx=5, pady=5)
            self.variables[var_name] = var

        # 設定欄寬
        self.columnconfigure(1, weight=1)

    def get_values(self):
        """獲取所有欄位值"""
        return {name: var.get() for name, var in self.variables.items()}

    def set_values(self, values):
        """設定所有欄位值"""
        for name, value in values.items():
            if name in self.variables:
                self.variables[name].set(value)


class FileSelector(ttk.Frame):
    """檔案選擇器元件"""

    def __init__(self, parent, label_text, button_text="瀏覽", file_types=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.file_types = file_types or [("All files", "*.*")]
        self.selected_path = None

        # 標籤
        ttk.Label(self, text=f"{label_text}:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)

        # 顯示欄位
        self.display_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.display_var, state="readonly").grid(row=0, column=1, sticky=tk.EW, padx=5, pady=5)

        # 瀏覽按鈕
        ttk.Button(self, text=button_text, command=self._browse_file, width=8).grid(row=0, column=2, sticky=tk.EW, padx=5, pady=5)

        # 設定欄寬
        self.columnconfigure(1, weight=1)

    def _browse_file(self):
        """瀏覽檔案"""
        file_path = filedialog.askopenfilename(filetypes=self.file_types)
        if file_path:
            self.selected_path = file_path
            self.display_var.set(os.path.basename(file_path))

    def get_path(self):
        """獲取選擇的路徑"""
        return self.selected_path

    def set_display(self, text):
        """設定顯示文字"""
        self.display_var.set(text)


class FolderFileSelector(ttk.Frame):
    """資料夾/檔案選擇器元件"""

    def __init__(self, parent, label_text, button_text="瀏覽", **kwargs):
        super().__init__(parent, **kwargs)
        self.selected_files = []
        self.mode_var = tk.StringVar(value="資料夾")

        # 模式選擇
        ttk.Label(self, text="模式:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Combobox(
            self,
            textvariable=self.mode_var,
            values=["資料夾", "資料"],
            state='readonly',
            width=8
        ).grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

        # 標籤和顯示欄位
        ttk.Label(self, text=f"{label_text}:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.display_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.display_var, state="readonly").grid(row=1, column=1, sticky=tk.EW, padx=5, pady=5)

        # 瀏覽按鈕
        ttk.Button(self, text=button_text, command=self._browse, width=8).grid(row=1, column=2, sticky=tk.EW, padx=5, pady=5)

        # 設定欄寬
        self.columnconfigure(1, weight=1)

    def _browse(self):
        """瀏覽檔案或資料夾"""
        mode = self.mode_var.get()
        if mode == "資料":
            file_paths = filedialog.askopenfilenames(
                title="選擇待測文件（可多選）",
                filetypes=[
                    ("PDF/圖片檔", "*.pdf *.jpeg *.jpg *.png *.bmp *.tif *.tiff"),
                    ("All files", "*.*")
                ]
            )
            if file_paths:
                self.selected_files = list(file_paths)
                self._update_display()
        else:
            folder_path = filedialog.askdirectory(title="選擇待測文件資料夾")
            if folder_path:
                # 掃描資料夾中的有效檔案
                valid_extensions = {'.pdf', '.jpeg', '.jpg', '.png', '.bmp', '.tif', '.tiff'}
                all_files = []
                for file_name in os.listdir(folder_path):
                    file_path = os.path.join(folder_path, file_name)
                    if os.path.isfile(file_path):
                        _, ext = os.path.splitext(file_name)
                        if ext.lower() in valid_extensions:
                            all_files.append(file_path)

                if all_files:
                    self.selected_files = all_files
                    self._update_display()
                else:
                    messagebox.showwarning(
                        "無有效檔案",
                        f"資料夾中沒有找到有效的 PDF 或圖片檔案\n\n支援格式: .pdf, .jpeg, .jpg, .png, .bmp, .tif, .tiff"
                    )

    def _update_display(self):
        """更新顯示文字"""
        if self.selected_files:
            filenames = ', '.join([os.path.basename(f) for f in self.selected_files[:3]])
            if len(self.selected_files) > 3:
                filenames += f" ... 等 {len(self.selected_files)} 個檔案"
            self.display_var.set(filenames)
        else:
            self.display_var.set("")

    def get_files(self):
        """獲取選擇的檔案列表"""
        return self.selected_files

    def set_display(self, text):
        """設定顯示文字"""
        self.display_var.set(text)


class LogDisplay(ttk.LabelFrame):
    """日誌顯示元件"""

    def __init__(self, parent, title="即時 Log 輸出", **kwargs):
        super().__init__(parent, text=title, **kwargs)

        # 建立 ScrolledText
        self.textbox = scrolledtext.ScrolledText(self, wrap=tk.WORD, font=("Consolas", 10))
        self.textbox.pack(fill=tk.BOTH, expand=True)

    def get_text_widget(self):
        """獲取文字元件（供 LoggerManager 使用）"""
        return self.textbox

    def clear(self):
        """清空日誌"""
        self.textbox.delete(1.0, tk.END)

    def append(self, text):
        """添加文字到日誌"""
        self.textbox.insert(tk.END, text + '\n')
        self.textbox.see(tk.END)


class StatusBar(ttk.Frame):
    """狀態列元件"""

    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)

        # 狀態標籤
        self.status_var = tk.StringVar(value="就緒")
        ttk.Label(self, textvariable=self.status_var).pack(side=tk.LEFT, padx=5)

        # 進度列
        self.progress = ttk.Progressbar(self, mode='determinate', length=200)
        self.progress.pack(side=tk.RIGHT, padx=5)

    def set_status(self, text):
        """設定狀態文字"""
        self.status_var.set(text)

    def set_progress(self, value):
        """設定進度值 (0-100)"""
        self.progress['value'] = value

    def start_progress(self):
        """開始不定進度"""
        self.progress.config(mode='indeterminate')
        self.progress.start()

    def stop_progress(self):
        """停止進度"""
        self.progress.stop()
        self.progress.config(mode='determinate')
        self.progress['value'] = 0
"""
ICR 辨識率測試系統 - Modern GUI 介面
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import configparser

from core.config import ConfigManager
from core.logger import LoggerManager
from core.orchestrator import TestOrchestrator

class ICRModernApp(tk.Tk):
    def __init__(self):
        super().__init__()
        
        # 設定 Windows 任務欄圖標（必須在其他設定之前）
        try:
            import ctypes
            myappid = 'TPI.ICR.PerformanceTest.1.0'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except:
            pass
        
        self.title("ICR 辨識率測試工具")
        self.geometry("700x750")
        self.minsize(700, 750)
        self.resizable(True, True)
        
        # 設定應用程式圖標
        try:
            if getattr(sys, 'frozen', False):
                # 打包後的環境
                base_path = sys._MEIPASS
            else:
                # 開發環境
                base_path = os.path.dirname(os.path.dirname(__file__))
            
            icon_path = os.path.join(base_path, 'lighting.ico')
            if os.path.exists(icon_path):
                # 設置窗口圖標
                self.iconbitmap(icon_path)
                # 設置任務欄圖標（Windows）
                self.wm_iconbitmap(default=icon_path)
        except Exception as e:
            print(f"無法載入圖標: {e}")
        
        # 創建選單欄
        self.create_menu_bar()
        
        # 配置管理器
        self.config_manager = ConfigManager()
        self.orchestrator = TestOrchestrator(self.config_manager)
        self.config_file = "config.ini"
        self.config = self.load_config()
        
        # UI 狀態
        self.temp_answer_file = None
        self.temp_upload_files = []
        self.selected_doc_type = '1'
        self.result_file_path = None
        self.is_running = False
        self.stop_requested = False
        self.test_thread = None
        self.stop_event = threading.Event()
        
        # 文件類型映射
        self.doc_type_map = {
            '居留證 (ARC)': '1',
            '體檢報告 (Health)': '2',
            '聘可函 (Employment)': '3'
        }
        
        # 初始化 GUI 變數
        self.doc_type_var = tk.StringVar(value='居留證 (ARC)')
        self.answer_file_var = tk.StringVar()
        self.answer_format_var = tk.StringVar(value="分行呈現")
        self.upload_mode_var = tk.StringVar(value="資料夾")
        self.upload_path_var = tk.StringVar()
        self.log_path_var = tk.StringVar(value=r"C:\Users\howie\Dev\TPI_Software\QA\Product\Systalk_ICR\Performance-Test\Log")
        
        # 保存答案形式控件的引用（用於動態顯示/隱藏）
        self.answer_format_label = None
        self.answer_format_combobox = None
        
        # 設定頁變數
        self.sftp_protocol_var = tk.StringVar(value="SFTP")
        self.sftp_host_var = tk.StringVar(value="192.168.160.67")
        self.sftp_port_var = tk.StringVar(value="22")
        self.sftp_username_var = tk.StringVar(value="tpiuser")
        self.sftp_password_var = tk.StringVar(value="1qaz@WSX3edc")
        self.sftp_remote_path_var = tk.StringVar(value="/home/tpiuser/icr-backend/imports/taipei/")
        
        self.db_host_var = tk.StringVar(value="192.168.160.67")
        self.db_port_var = tk.StringVar(value="5555")
        self.db_database_var = tk.StringVar(value="postgres")
        self.db_username_var = tk.StringVar(value="postgres")
        self.db_password_var = tk.StringVar(value="1qaz@WSX3edc")
        
        # 建立 UI
        self.setup_ui()
        
        # 設定日誌系統
        self.setup_logging()
        
        # 載入配置到 UI
        self.load_config_to_ui()
        
        # 更新日誌路徑（如果有變化）
        self.update_logging_path()
    
    def create_menu_bar(self):
        """創建選單欄"""
        menu_bar = tk.Menu(self)
        self.config(menu=menu_bar)
        
        # Action 選單
        action_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Action", menu=action_menu)
        action_menu.add_command(label="Run", command=self.start_testing)
        action_menu.add_command(label="Stop", command=self.stop_testing)
        
        # Config 選單
        config_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Config", menu=config_menu)
        config_menu.add_command(label="Load", command=self.load_config_from_file)
        config_menu.add_command(label="Save", command=self.save_config)
        config_menu.add_command(label="Save As", command=self.save_config_as)
    
    def load_config(self):
        """載入配置文件"""
        config = configparser.ConfigParser()
        if os.path.exists(self.config_file):
            config.read(self.config_file, encoding='utf-8')
        return config

    def setup_logging(self):
        """設定日誌系統"""
        log_dir = self.config_manager.paths.get_log_dir()
        self.logger = LoggerManager.setup_logger(
            log_dir=log_dir,
            text_widget=self.log_textbox
        )
    
    def update_logging_path(self):
        """更新日誌路徑"""
        log_path = self.log_path_var.get()
        if log_path and log_path != self.config_manager.paths.get_log_dir():
            self.config_manager.set_log_dir(log_path)
            # 如果 logger 已存在，需要重新設置
            if hasattr(self, 'logger') and self.logger:
                # 重新設置 logger 到新路徑
                log_dir = self.config_manager.paths.get_log_dir()
                self.logger = LoggerManager.setup_logger(
                    log_dir=log_dir,
                    text_widget=self.log_textbox
                )

    def setup_ui(self):
        """建立現代化 UI 介面"""
        # ===== 按鈕區塊 (選單下方) =====
        run_btn_frame = ttk.Frame(self)
        run_btn_frame.pack(fill=tk.X, padx=0, pady=(0, 0))
        
        self.run_btn = ttk.Button(run_btn_frame, text="🚀 Run", width=7, command=self.start_testing)
        self.run_btn.pack(side="left", padx=(4, 2), pady=0)
        
        self.stop_btn = ttk.Button(run_btn_frame, text="⛔ Stop", width=7, command=self.stop_testing, state="disabled")
        self.stop_btn.pack(side="left", padx=(2, 2), pady=0)
        
        self.no_answer_run_btn = ttk.Button(run_btn_frame, text="📤 No Answer Run", width=15, command=self.start_no_answer_testing)
        self.no_answer_run_btn.pack(side="left", padx=(2, 0), pady=0)
        
        self.open_result_btn = ttk.Button(run_btn_frame, text="📄 結果", width=8, command=self.open_result_file, state="disabled")
        self.open_result_btn.pack(side="right", padx=(4, 8), pady=0)
        
        # ===== 上方：分頁區域 =====
        tab_frame = ttk.Frame(self)
        tab_frame.pack(fill=tk.BOTH, expand=False, padx=2, pady=2)
        
        self.notebook = ttk.Notebook(tab_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # 創建執行頁
        self.create_execution_tab()
        
        # 創建設定頁
        self.create_sftp_tab()
        self.create_database_tab()
        
        # ===== 下方：Log 輸出區域 (固定高度) =====
        log_frame = ttk.LabelFrame(self, text="即時 Log 輸出", padding="2")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        
        self.log_textbox = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, font=("Consolas", 10), height=10)
        self.log_textbox.pack(fill=tk.BOTH, expand=True)
    
    def create_execution_tab(self):
        """創建執行頁分頁"""
        exec_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(exec_tab, text="執行頁")
        
        # ===== 第一步驟：文件類型 =====
        step1_frame = ttk.LabelFrame(exec_tab, text="第一步驟：選擇文件類型", padding="5")
        step1_frame.pack(fill=tk.X, pady=5)
        step1_frame.columnconfigure(1, weight=1)
        
        ttk.Label(step1_frame, text="文件類型:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Combobox(
            step1_frame,
            textvariable=self.doc_type_var,
            values=list(self.doc_type_map.keys()),
            state='readonly',
            width=25
        ).grid(row=0, column=1, sticky=tk.EW, padx=5, pady=5)
        
        self.doc_type_var.trace("w", lambda *args: self.on_doc_type_changed())
        
        # ===== 第二步驟：上傳答案 =====
        step2_frame = ttk.LabelFrame(exec_tab, text="第二步驟：上傳答案檔案 (xlsx/csv)", padding="5")
        step2_frame.pack(fill=tk.X, pady=5)
        
        # 調整權重：現在要讓路徑 Entry (原本在 col 1，現在移到 col 4) 能夠伸縮
        step2_frame.columnconfigure(4, weight=1)
        step2_frame.columnconfigure(1, weight=0) # 重置原本的權重

        # --- 左側：答案形式（只在聘可函時啟用）---
        self.answer_format_label = ttk.Label(step2_frame, text="答案形式:")
        self.answer_format_label.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.answer_format_combobox = ttk.Combobox(
            step2_frame,
            textvariable=self.answer_format_var,
            values=["分行呈現", "列表呈現"],
            state="disabled",  # 初始時反灰（預設是 ARC）
            width=12
        )
        self.answer_format_combobox.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

        # --- 右側：答案檔案路徑 ---
        ttk.Label(step2_frame, text="答案:").grid(row=0, column=2, sticky=tk.W, padx=(15, 5), pady=5)
        ttk.Entry(step2_frame, textvariable=self.answer_file_var, state="readonly").grid(row=0, column=3, columnspan=2, sticky=tk.EW, padx=5, pady=5)
        ttk.Button(step2_frame, text="瀏覽", command=self.upload_answer_file, width=8).grid(row=0, column=5, sticky=tk.EW, padx=5, pady=5)
        
        # ===== 第三步驟：上傳待測文件 =====
        step3_frame = ttk.LabelFrame(exec_tab, text="第三步驟：上傳待測文件", padding="5")
        step3_frame.pack(fill=tk.X, pady=5)
        step3_frame.columnconfigure(2, weight=1)
        
        ttk.Label(step3_frame, text="模式:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Combobox(
            step3_frame,
            textvariable=self.upload_mode_var,
            values=["資料夾", "資料"],
            state='readonly',
            width=8
        ).grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(step3_frame, textvariable=self.upload_path_var, state="readonly").grid(row=0, column=2, sticky=tk.EW, padx=5, pady=5)
        ttk.Button(step3_frame, text="瀏覽", command=self.upload_test, width=8).grid(row=0, column=3, sticky=tk.EW, padx=5, pady=5)
        
        # ===== 第四步驟：指定 Log 路徑 =====
        step4_frame = ttk.LabelFrame(exec_tab, text="第四步驟：指定 Log 路徑", padding="5")
        step4_frame.pack(fill=tk.X, pady=5)
        step4_frame.columnconfigure(1, weight=1)
        
        ttk.Label(step4_frame, text="Log 路徑:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(step4_frame, textvariable=self.log_path_var, state="readonly").grid(row=0, column=1, sticky=tk.EW, padx=5, pady=5)
        ttk.Button(step4_frame, text="瀏覽", command=self.select_log_path, width=8).grid(row=0, column=2, sticky=tk.EW, padx=5, pady=5)
    
    def create_sftp_tab(self):
        """創建 SFTP 設定頁分頁"""
        sftp_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(sftp_tab, text="WinSCP")
        
        # ===== SFTP 設定 =====
        sftp_frame = ttk.LabelFrame(sftp_tab, text="WinSCP 設定", padding="10")
        sftp_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Protocol
        ttk.Label(sftp_frame, text="Protocol:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(sftp_frame, textvariable=self.sftp_protocol_var).grid(row=0, column=1, sticky=tk.EW, padx=5, pady=5)
        
        # Host
        ttk.Label(sftp_frame, text="Host:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(sftp_frame, textvariable=self.sftp_host_var).grid(row=1, column=1, sticky=tk.EW, padx=5, pady=5)
        
        # Port
        ttk.Label(sftp_frame, text="Port:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(sftp_frame, textvariable=self.sftp_port_var).grid(row=2, column=1, sticky=tk.EW, padx=5, pady=5)
        
        # Username
        ttk.Label(sftp_frame, text="Username:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(sftp_frame, textvariable=self.sftp_username_var).grid(row=3, column=1, sticky=tk.EW, padx=5, pady=5)
        
        # Password
        ttk.Label(sftp_frame, text="Password:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(sftp_frame, textvariable=self.sftp_password_var, show="*").grid(row=4, column=1, sticky=tk.EW, padx=5, pady=5)
        
        # Remote Path
        ttk.Label(sftp_frame, text="Remote Path:").grid(row=5, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(sftp_frame, textvariable=self.sftp_remote_path_var).grid(row=5, column=1, sticky=tk.EW, padx=5, pady=5)
        
        # 設定欄寬
        sftp_frame.columnconfigure(1, weight=1)
    
    def create_database_tab(self):
        """創建 DATABASE 設定頁分頁"""
        db_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(db_tab, text="DATABASE")
        
        # ===== DATABASE 設定 =====
        db_frame = ttk.LabelFrame(db_tab, text="DATABASE 設定", padding="10")
        db_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Host
        ttk.Label(db_frame, text="Host:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(db_frame, textvariable=self.db_host_var).grid(row=0, column=1, sticky=tk.EW, padx=5, pady=5)
        
        # Port
        ttk.Label(db_frame, text="Port:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(db_frame, textvariable=self.db_port_var).grid(row=1, column=1, sticky=tk.EW, padx=5, pady=5)
        
        # Database
        ttk.Label(db_frame, text="Database:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(db_frame, textvariable=self.db_database_var).grid(row=2, column=1, sticky=tk.EW, padx=5, pady=5)
        
        # Username
        ttk.Label(db_frame, text="Username:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(db_frame, textvariable=self.db_username_var).grid(row=3, column=1, sticky=tk.EW, padx=5, pady=5)
        
        # Password
        ttk.Label(db_frame, text="Password:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(db_frame, textvariable=self.db_password_var, show="*").grid(row=4, column=1, sticky=tk.EW, padx=5, pady=5)
        
        # 設定欄寬
        db_frame.columnconfigure(1, weight=1)
    
    def setup_logging(self):
        """設定日誌系統"""
        log_dir = self.log_path_var.get()
        self.logger = LoggerManager.setup_logger(
            log_dir=log_dir,
            text_widget=self.log_textbox
        )
    
    def on_doc_type_changed(self):
        """文件類型變更事件"""
        selected = self.doc_type_var.get()
        self.selected_doc_type = self.doc_type_map[selected]
        self.logger.info(f"已選擇文件類型: {selected}")
        
        # 根據選擇的文件類型啟用/停用答案形式控件
        if self.answer_format_combobox:
            if selected == '聘可函 (Employment)':
                # 啟用答案形式控件
                self.answer_format_combobox.config(state="readonly")
            else:
                # 停用答案形式控件（反灰）
                self.answer_format_combobox.config(state="disabled")
    
    def upload_answer_file(self):
        """上傳答案檔案"""
        file_path = filedialog.askopenfilename(
            title="選擇答案檔案",
            filetypes=[("Excel/CSV files", "*.xlsx *.csv"), ("All files", "*.*")]
        )
        if file_path:
            self.temp_answer_file = file_path
            filename = os.path.basename(file_path)
            self.answer_file_var.set(filename)
            self.logger.info(f"已選擇答案檔案: {filename}")
            # 立即讀取並顯示答案檔案內容
            try:
                import pandas as pd
                if file_path.endswith('.xlsx') or file_path.endswith('.xls'):
                    df = pd.read_excel(file_path)
                elif file_path.endswith('.csv'):
                    df = pd.read_csv(file_path)
                else:
                    self.logger.info("[答案檔案內容] 不支援的檔案格式")
                    return
                self.logger.info("[答案檔案內容] 檔案共 %d 筆" % len(df))
                # 顯示所有檔名欄位
                col_candidates = [c for c in df.columns if '檔名' in c or 'filename' in c.lower() or 'file' in c.lower()]
                if col_candidates:
                    col = col_candidates[0]
                    filenames = df[col].astype(str).tolist()
                    for i, fname in enumerate(filenames, 1):
                        self.logger.info(f"  [{i}] {fname}")
                else:
                    self.logger.info("[答案檔案內容] 無檔名欄位，欄位: " + ', '.join(df.columns))
            except Exception as e:
                self.logger.warning(f"讀取答案檔案內容失敗: {e}")
    
    def upload_test(self):
        """上傳待測文件"""
        mode = self.upload_mode_var.get()
        if mode == "資料":
            file_paths = filedialog.askopenfilenames(
                title="選擇待測文件（可多選）",
                filetypes=[
                    ("PDF/圖片檔", "*.pdf *.jpeg *.jpg *.png *.bmp *.tif *.tiff"),
                    ("All files", "*.*")
                ]
            )
            if file_paths:
                self.temp_upload_files = list(file_paths)
                filenames = ', '.join([os.path.basename(f) for f in self.temp_upload_files[:3]])
                if len(self.temp_upload_files) > 3:
                    filenames += f" ... 等 {len(self.temp_upload_files)} 個檔案"
                self.upload_path_var.set(filenames)
                self.logger.info(f"已選擇 {len(self.temp_upload_files)} 個待測文件")
                self.logger.info("待測文件清單：")
                for i, f in enumerate(self.temp_upload_files, 1):
                    self.logger.info(f"  [{i}] {os.path.basename(f)}")
        else:
            folder_path = filedialog.askdirectory(title="選擇待測文件資料夾")
            if folder_path:
                valid_extensions = {'.pdf', '.jpeg', '.jpg', '.png', '.bmp', '.tif', '.tiff'}
                all_files = []
                for file_name in os.listdir(folder_path):
                    file_path = os.path.join(folder_path, file_name)
                    if os.path.isfile(file_path):
                        _, ext = os.path.splitext(file_name)
                        if ext.lower() in valid_extensions:
                            all_files.append(file_path)
                if all_files:
                    self.temp_upload_files = all_files
                    filenames = ', '.join([os.path.basename(f) for f in all_files[:3]])
                    if len(all_files) > 3:
                        filenames += f" ... 等 {len(all_files)} 個檔案"
                    self.upload_path_var.set(filenames)
                    folder_name = os.path.basename(folder_path)
                    self.logger.info(f"從資料夾選擇: {folder_name}")
                    self.logger.info(f"已選擇 {len(all_files)} 個待測文件")
                    self.logger.info("待測文件清單：")
                    for i, f in enumerate(all_files, 1):
                        self.logger.info(f"  [{i}] {os.path.basename(f)}")
                else:
                    self.logger.warning(f"資料夾中沒有找到有效的 PDF 或圖片檔案: {folder_path}")
                    messagebox.showwarning(
                        "無有效檔案",
                        f"資料夾中沒有找到有效的 PDF 或圖片檔案\n\n支援格式: .pdf, .jpeg, .jpg, .png, .bmp, .tif, .tiff"
                    )
    
    def select_log_path(self):
        """選擇 Log 路徑"""
        folder_path = filedialog.askdirectory(title="選擇 Log 路徑")
        if folder_path:
            self.log_path_var.set(folder_path)
            self.config_manager.set_log_dir(folder_path)
            self.logger.info(f"已選擇 Log 路徑: {folder_path}")
            # 重新設定日誌系統到新路徑
            self.setup_logging()
    
    def load_config_to_ui(self):
        """從配置文件載入數據到 UI"""
        try:
            if self.config.has_section('testing'):
                if self.config.has_option('testing', 'doc_type'):
                    doc_type_name = self.config.get('testing', 'doc_type')
                    self.doc_type_var.set(doc_type_name)
                if self.config.has_option('testing', 'answer_format'):
                    answer_format = self.config.get('testing', 'answer_format')
                    self.answer_format_var.set(answer_format)
            
            if self.config.has_section('files'):
                if self.config.has_option('files', 'answer_file'):
                    answer_file = self.config.get('files', 'answer_file')
                    if answer_file and os.path.exists(answer_file):
                        self.temp_answer_file = answer_file
                        self.answer_file_var.set(os.path.basename(answer_file))
                    else:
                        self.answer_file_var.set(answer_file)
                if self.config.has_option('files', 'upload_path'):
                    self.upload_path_var.set(self.config.get('files', 'upload_path'))
                if self.config.has_option('files', 'upload_files'):
                    upload_files_str = self.config.get('files', 'upload_files')
                    if upload_files_str:
                        import json
                        try:
                            self.temp_upload_files = json.loads(upload_files_str)
                        except:
                            pass
            
            # 載入 SFTP 設定
            if self.config.has_section('SFTP'):
                if self.config.has_option('SFTP', 'protocol'):
                    self.sftp_protocol_var.set(self.config.get('SFTP', 'protocol'))
                if self.config.has_option('SFTP', 'host'):
                    self.sftp_host_var.set(self.config.get('SFTP', 'host'))
                if self.config.has_option('SFTP', 'port'):
                    self.sftp_port_var.set(self.config.get('SFTP', 'port'))
                if self.config.has_option('SFTP', 'username'):
                    self.sftp_username_var.set(self.config.get('SFTP', 'username'))
                if self.config.has_option('SFTP', 'password'):
                    self.sftp_password_var.set(self.config.get('SFTP', 'password'))
                if self.config.has_option('SFTP', 'remote_path'):
                    self.sftp_remote_path_var.set(self.config.get('SFTP', 'remote_path'))
            
            # 載入 DATABASE 設定
            if self.config.has_section('DATABASE'):
                if self.config.has_option('DATABASE', 'host'):
                    self.db_host_var.set(self.config.get('DATABASE', 'host'))
                if self.config.has_option('DATABASE', 'port'):
                    self.db_port_var.set(self.config.get('DATABASE', 'port'))
                if self.config.has_option('DATABASE', 'database'):
                    self.db_database_var.set(self.config.get('DATABASE', 'database'))
                if self.config.has_option('DATABASE', 'user'):
                    self.db_username_var.set(self.config.get('DATABASE', 'user'))
                if self.config.has_option('DATABASE', 'password'):
                    self.db_password_var.set(self.config.get('DATABASE', 'password'))
            
            # 載入 Path 設定
            if self.config.has_section('Path'):
                if self.config.has_option('Path', 'log_path'):
                    log_path = self.config.get('Path', 'log_path')
                    self.log_path_var.set(log_path)
                    self.config_manager.set_log_dir(log_path)
                    
        except Exception as e:
            self.logger.warning(f"載入配置時發生錯誤: {e}")
    
    def save_config(self):
        """儲存配置到文件"""
        try:
            if not self.config.has_section('testing'):
                self.config.add_section('testing')
            self.config.set('testing', 'doc_type', self.doc_type_var.get())
            self.config.set('testing', 'answer_format', self.answer_format_var.get())
            
            if not self.config.has_section('files'):
                self.config.add_section('files')
            # 儲存實際的文件路徑而非顯示名稱
            self.config.set('files', 'answer_file', self.temp_answer_file if self.temp_answer_file else '')
            self.config.set('files', 'upload_path', self.upload_path_var.get())
            # 儲存上傳文件列表
            import json
            self.config.set('files', 'upload_files', json.dumps(self.temp_upload_files, ensure_ascii=False) if self.temp_upload_files else '')
            
            # 儲存 SFTP 設定
            if not self.config.has_section('SFTP'):
                self.config.add_section('SFTP')
            self.config.set('SFTP', 'protocol', self.sftp_protocol_var.get())
            self.config.set('SFTP', 'host', self.sftp_host_var.get())
            self.config.set('SFTP', 'port', self.sftp_port_var.get())
            self.config.set('SFTP', 'username', self.sftp_username_var.get())
            self.config.set('SFTP', 'password', self.sftp_password_var.get())
            self.config.set('SFTP', 'remote_path', self.sftp_remote_path_var.get())
            
            # 儲存 DATABASE 設定
            if not self.config.has_section('DATABASE'):
                self.config.add_section('DATABASE')
            self.config.set('DATABASE', 'host', self.db_host_var.get())
            self.config.set('DATABASE', 'port', self.db_port_var.get())
            self.config.set('DATABASE', 'database', self.db_database_var.get())
            self.config.set('DATABASE', 'user', self.db_username_var.get())
            self.config.set('DATABASE', 'password', self.db_password_var.get())
            
            # 儲存 Path 設定
            if not self.config.has_section('Path'):
                self.config.add_section('Path')
            self.config.set('Path', 'log_path', self.log_path_var.get())
            
            with open(self.config_file, 'w', encoding='utf-8') as f:
                self.config.write(f)
            
            self.logger.info(f"配置已儲存到 {self.config_file}")
            messagebox.showinfo("成功", f"配置已儲存到 {self.config_file}")
        except Exception as e:
            error_msg = f"儲存配置失敗: {e}"
            self.logger.error(error_msg)
            messagebox.showerror("錯誤", error_msg)
    
    def save_config_as(self):
        """另存配置文件"""
        try:
            file_path = filedialog.asksaveasfilename(
                title="另存配置文件",
                defaultextension=".ini",
                filetypes=[("INI Files", "*.ini"), ("All Files", "*.*")]
            )
            if file_path:
                if not self.config.has_section('testing'):
                    self.config.add_section('testing')
                self.config.set('testing', 'doc_type', self.doc_type_var.get())
                self.config.set('testing', 'answer_format', self.answer_format_var.get())
                
                if not self.config.has_section('files'):
                    self.config.add_section('files')
                # 儲存實際的文件路徑而非顯示名稱
                self.config.set('files', 'answer_file', self.temp_answer_file if self.temp_answer_file else '')
                self.config.set('files', 'upload_path', self.upload_path_var.get())
                # 儲存上傳文件列表
                import json
                self.config.set('files', 'upload_files', json.dumps(self.temp_upload_files, ensure_ascii=False) if self.temp_upload_files else '')
                
                # 儲存 SFTP 設定
                if not self.config.has_section('SFTP'):
                    self.config.add_section('SFTP')
                self.config.set('SFTP', 'protocol', self.sftp_protocol_var.get())
                self.config.set('SFTP', 'host', self.sftp_host_var.get())
                self.config.set('SFTP', 'port', self.sftp_port_var.get())
                self.config.set('SFTP', 'username', self.sftp_username_var.get())
                self.config.set('SFTP', 'password', self.sftp_password_var.get())
                self.config.set('SFTP', 'remote_path', self.sftp_remote_path_var.get())
                
                # 儲存 DATABASE 設定
                if not self.config.has_section('DATABASE'):
                    self.config.add_section('DATABASE')
                self.config.set('DATABASE', 'host', self.db_host_var.get())
                self.config.set('DATABASE', 'port', self.db_port_var.get())
                self.config.set('DATABASE', 'database', self.db_database_var.get())
                self.config.set('DATABASE', 'user', self.db_username_var.get())
                self.config.set('DATABASE', 'password', self.db_password_var.get())
                
                # 儲存 Path 設定
                if not self.config.has_section('Path'):
                    self.config.add_section('Path')
                self.config.set('Path', 'log_path', self.log_path_var.get())
                
                with open(file_path, 'w', encoding='utf-8') as f:
                    self.config.write(f)
                
                self.logger.info(f"配置已另存到 {file_path}")
                messagebox.showinfo("成功", f"配置已另存到 {os.path.basename(file_path)}")
        except Exception as e:
            error_msg = f"另存配置失敗: {e}"
            self.logger.error(error_msg)
            messagebox.showerror("錯誤", error_msg)
    
    def load_config_from_file(self):
        """從文件載入配置"""
        try:
            file_path = filedialog.askopenfilename(
                title="選擇配置文件",
                defaultextension=".ini",
                filetypes=[("INI Files", "*.ini"), ("All Files", "*.*")]
            )
            if file_path:
                new_config = configparser.ConfigParser()
                new_config.read(file_path, encoding='utf-8')
                
                if new_config.has_section('testing'):
                    if new_config.has_option('testing', 'doc_type'):
                        self.doc_type_var.set(new_config.get('testing', 'doc_type'))
                    if new_config.has_option('testing', 'answer_format'):
                        self.answer_format_var.set(new_config.get('testing', 'answer_format'))
                
                if new_config.has_section('files'):
                    if new_config.has_option('files', 'answer_file'):
                        answer_file = new_config.get('files', 'answer_file')
                        if answer_file and os.path.exists(answer_file):
                            self.temp_answer_file = answer_file
                            self.answer_file_var.set(os.path.basename(answer_file))
                        else:
                            self.answer_file_var.set(answer_file)
                    if new_config.has_option('files', 'upload_path'):
                        self.upload_path_var.set(new_config.get('files', 'upload_path'))
                    if new_config.has_option('files', 'upload_files'):
                        upload_files_str = new_config.get('files', 'upload_files')
                        if upload_files_str:
                            import json
                            try:
                                self.temp_upload_files = json.loads(upload_files_str)
                            except:
                                pass
                
                # 載入 SFTP 設定
                if new_config.has_section('SFTP'):
                    if new_config.has_option('SFTP', 'protocol'):
                        self.sftp_protocol_var.set(new_config.get('SFTP', 'protocol'))
                    if new_config.has_option('SFTP', 'host'):
                        self.sftp_host_var.set(new_config.get('SFTP', 'host'))
                    if new_config.has_option('SFTP', 'port'):
                        self.sftp_port_var.set(new_config.get('SFTP', 'port'))
                    if new_config.has_option('SFTP', 'username'):
                        self.sftp_username_var.set(new_config.get('SFTP', 'username'))
                    if new_config.has_option('SFTP', 'password'):
                        self.sftp_password_var.set(new_config.get('SFTP', 'password'))
                    if new_config.has_option('SFTP', 'remote_path'):
                        self.sftp_remote_path_var.set(new_config.get('SFTP', 'remote_path'))
                
                # 載入 DATABASE 設定
                if new_config.has_section('DATABASE'):
                    if new_config.has_option('DATABASE', 'host'):
                        self.db_host_var.set(new_config.get('DATABASE', 'host'))
                    if new_config.has_option('DATABASE', 'port'):
                        self.db_port_var.set(new_config.get('DATABASE', 'port'))
                    if new_config.has_option('DATABASE', 'database'):
                        self.db_database_var.set(new_config.get('DATABASE', 'database'))
                    if new_config.has_option('DATABASE', 'user'):
                        self.db_username_var.set(new_config.get('DATABASE', 'user'))
                    if new_config.has_option('DATABASE', 'password'):
                        self.db_password_var.set(new_config.get('DATABASE', 'password'))
                
                # 載入 Path 設定
                if new_config.has_section('Path'):
                    if new_config.has_option('Path', 'log_path'):
                        log_path = new_config.get('Path', 'log_path')
                        self.log_path_var.set(log_path)
                        self.config_manager.set_log_dir(log_path)
                
                self.config = new_config
                self.config_file = file_path
                self.logger.info(f"配置已從 {os.path.basename(file_path)} 載入")
                messagebox.showinfo("成功", f"配置已從 {os.path.basename(file_path)} 載入")
        except Exception as e:
            error_msg = f"載入配置失敗: {e}"
            self.logger.error(error_msg)
            messagebox.showerror("錯誤", error_msg)
    
    def start_testing(self):
        """開始評分流程"""
        if self.is_running:
            messagebox.showwarning("警告", "測試正在執行中，請等待完成或強制終止")
            return
        
        if not self.temp_answer_file or not self.temp_upload_files:
            messagebox.showwarning("警告", "請先選擇答案檔案和待測文件")
            return
        
        self.is_running = True
        self.stop_requested = False
        self.stop_event.clear()
        self.run_btn.config(state="disabled", text="執行中...")
        self.no_answer_run_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        
        # 禁用所有輸入欄位
        self._disable_input_fields()
        
        # 在新執行緒中執行測試流程
        self.test_thread = threading.Thread(target=self.run_test_thread, daemon=True)
        self.test_thread.start()
    
    def start_no_answer_testing(self):
        """開始無答案測試流程（只上傳和抓取DB）"""
        if self.is_running:
            messagebox.showwarning("警告", "測試正在執行中，請等待完成或強制終止")
            return
        
        if not self.temp_upload_files:
            messagebox.showwarning("警告", "請先選擇待測文件")
            return
        
        self.is_running = True
        self.stop_requested = False
        self.stop_event.clear()
        self.run_btn.config(state="disabled")
        self.no_answer_run_btn.config(state="disabled", text="執行中...")
        self.stop_btn.config(state="normal")
        
        # 禁用所有輸入欄位
        self._disable_input_fields()
        
        # 在新執行緒中執行測試流程
        self.test_thread = threading.Thread(target=self.run_no_answer_test_thread, daemon=True)
        self.test_thread.start()
    
    def run_test_thread(self):
        """執行測試流程（在獨立執行緒中）"""
        try:
            # 收集當前 GUI 設定作為覆蓋參數
            sftp_config_override = {
                'host': self.sftp_host_var.get(),
                'port': self.sftp_port_var.get(),
                'username': self.sftp_username_var.get(),
                'password': self.sftp_password_var.get(),
                'remote_path': self.sftp_remote_path_var.get()
            }
            
            db_config_override = {
                'host': self.db_host_var.get(),
                'port': self.db_port_var.get(),
                'database': self.db_database_var.get(),
                'user': self.db_username_var.get(),
                'password': self.db_password_var.get()
            }
            
            result = self.orchestrator.execute_test_workflow(
                doc_type=self.selected_doc_type,
                answer_file_path=self.temp_answer_file,
                upload_files=self.temp_upload_files,
                stop_check_callback=lambda: self.stop_requested,
                sftp_config_override=sftp_config_override,
                db_config_override=db_config_override,
                answer_format=self.answer_format_var.get()
            )
            
            if result['success']:
                self.after(0, lambda: self._show_results(result['result_path'], result['statistics']))
                self.after(0, lambda: messagebox.showinfo(
                    "完成",
                    f"測試完成！\n結果已儲存至：{result['result_path']}"
                ))
        except Exception as e:
            self.logger.error(f"執行失敗: {e}")
            error_msg = str(e)
            self.after(0, lambda msg=error_msg: messagebox.showerror("錯誤", f"執行失敗：{msg}"))
        finally:
            self.is_running = False
            self.stop_requested = False
            self.after(0, self._reset_buttons)
    
    def run_no_answer_test_thread(self):
        """執行無答案測試流程（在獨立執行緒中）"""
        try:
            # 收集當前 GUI 設定作為覆蓋參數
            sftp_config_override = {
                'host': self.sftp_host_var.get(),
                'port': self.sftp_port_var.get(),
                'username': self.sftp_username_var.get(),
                'password': self.sftp_password_var.get(),
                'remote_path': self.sftp_remote_path_var.get()
            }
            
            db_config_override = {
                'host': self.db_host_var.get(),
                'port': self.db_port_var.get(),
                'database': self.db_database_var.get(),
                'user': self.db_username_var.get(),
                'password': self.db_password_var.get()
            }
            
            result = self.orchestrator.execute_no_answer_workflow(
                doc_type=self.selected_doc_type,
                upload_files=self.temp_upload_files,
                stop_check_callback=lambda: self.stop_requested,
                sftp_config_override=sftp_config_override,
                db_config_override=db_config_override,
                answer_format=self.answer_format_var.get()
            )
            
            if result['success']:
                self.after(0, lambda: self._show_no_answer_results(result['result_path']))
                self.after(0, lambda: messagebox.showinfo(
                    "完成",
                    f"無答案測試完成！\n結果已儲存至：{result['result_path']}"
                ))
        except Exception as e:
            self.logger.error(f"執行失敗: {e}")
            error_msg = str(e)
            self.after(0, lambda msg=error_msg: messagebox.showerror("錯誤", f"執行失敗：{msg}"))
        finally:
            self.is_running = False
            self.stop_requested = False
            self.after(0, self._reset_buttons)
    
    def _disable_input_fields(self):
        """禁用所有輸入欄位"""
        # 禁用執行頁的所有元件（不包括整個 notebook，因為它不支援 state 選項）
        for child in self.notebook.winfo_children():
            if hasattr(child, 'winfo_children'):
                for widget in child.winfo_children():
                    self._disable_widget_recursive(widget)
    
    def _enable_input_fields(self):
        """重新啟用所有輸入欄位"""
        # 重新啟用執行頁的所有元件（不包括整個 notebook，因為它不支援 state 選項）
        for child in self.notebook.winfo_children():
            if hasattr(child, 'winfo_children'):
                for widget in child.winfo_children():
                    self._enable_widget_recursive(widget)
    
    def _disable_widget_recursive(self, widget):
        """遞歸禁用元件"""
        try:
            if isinstance(widget, (ttk.Combobox, ttk.Entry, ttk.Button)):
                if widget != self.stop_btn:  # 除了 Stop 按鈕
                    widget.config(state="disabled")
            elif hasattr(widget, 'winfo_children'):
                for child in widget.winfo_children():
                    self._disable_widget_recursive(child)
        except:
            pass  # 忽略無法禁用的元件
    
    def _enable_widget_recursive(self, widget):
        """遞歸重新啟用元件"""
        try:
            if isinstance(widget, ttk.Combobox):
                widget.config(state="readonly")  # Combobox 恢復為 readonly
            elif isinstance(widget, ttk.Entry):
                widget.config(state="normal")  # Entry 恢復為 normal
            elif isinstance(widget, ttk.Button):
                widget.config(state="normal")  # Button 恢復為 normal
            elif hasattr(widget, 'winfo_children'):
                for child in widget.winfo_children():
                    self._enable_widget_recursive(child)
        except:
            pass  # 忽略無法重新啟用的元件
    
    def _reset_buttons(self):
        """重設按鈕狀態"""
        self.run_btn.config(state="normal", text="🚀 Run")
        self.no_answer_run_btn.config(state="normal", text="📤 No Answer Run")
        self.stop_btn.config(state="disabled", text="⛔ Stop")
        # 重新啟用所有輸入欄位
        self._enable_input_fields()
    
    def stop_testing(self):
        """強制終止測試流程"""
        if self.is_running:
            result = messagebox.askyesno("確認", "確定要強制終止測試嗎？")
            if result:
                self.stop_requested = True
                self.logger.warning("使用者請求強制終止...")
                self.stop_btn.config(state="disabled", text="終止中...")
    
    def _show_results(self, output_path: str, stats: dict):
        """顯示測試結果"""
        self.result_file_path = output_path
        self.open_result_btn.config(state="normal")
        self.logger.info(f"評分結果: PASS {stats['pass']} / FAIL {stats['fail']}")
    
    def _show_no_answer_results(self, output_path: str):
        """顯示無答案測試結果"""
        self.result_file_path = output_path
        self.open_result_btn.config(state="normal")
        self.logger.info(f"辨識結果已匯出")
    
    def open_result_file(self):
        """開啟結果檔案"""
        if self.result_file_path and os.path.exists(self.result_file_path):
            os.startfile(self.result_file_path)
        else:
            messagebox.showerror("錯誤", "結果檔案不存在")


if __name__ == "__main__":
    app = ICRModernApp()
    app.mainloop()

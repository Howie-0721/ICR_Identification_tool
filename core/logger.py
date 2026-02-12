"""
日誌管理模組
提供統一的日誌配置與管理功能
"""

import os
import logging
import datetime
from typing import Optional
import tkinter as tk
from tkinter import scrolledtext


class TextHandler(logging.Handler):
    """GUI 文字元件的日誌處理器"""
    
    def __init__(self, text_widget: scrolledtext.ScrolledText):
        logging.Handler.__init__(self)
        self.text_widget = text_widget
    
    def emit(self, record):
        """發送日誌訊息到文字元件"""
        msg = self.format(record)
        
        def append():
            self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.see(tk.END)
        
        self.text_widget.after(0, append)


class LoggerManager:
    """
    日誌管理器
    
    提供統一的日誌配置功能：
    - 主控台輸出
    - 檔案輸出
    - GUI 文字元件輸出
    """
    
    LOGGER_NAME = "ICRLogger"
    
    # 儲存當前的時間戳和 log 檔案路徑
    _current_timestamp: Optional[str] = None
    _current_log_file: Optional[str] = None
    
    @staticmethod
    def setup_logger(
        log_dir: str,
        level: int = logging.INFO,
        text_widget: Optional[scrolledtext.ScrolledText] = None
    ) -> logging.Logger:
        """
        設定日誌記錄器
        
        Args:
            log_dir: 日誌檔案目錄
            level: 日誌級別 (預設 INFO)
            text_widget: GUI 文字元件 (可選)
            
        Returns:
            配置好的 Logger 物件
        """
        logger = logging.getLogger(LoggerManager.LOGGER_NAME)
        logger.setLevel(level)
        logger.handlers.clear()  # 清除既有的處理器
        
        # 格式器
        formatter = logging.Formatter(
            '%(asctime)s [%(levelname)s] [%(filename)s]: %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        file_formatter = logging.Formatter(
            '%(asctime)s [%(levelname)s] %(filename)s: %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # 1. GUI 文字元件處理器 (如果提供)
        if text_widget is not None:
            text_handler = TextHandler(text_widget)
            text_handler.setFormatter(formatter)
            logger.addHandler(text_handler)
        
        # 2. 檔案處理器 - 直接寫入到 Log/<timestamp>/Log.txt
        os.makedirs(log_dir, exist_ok=True)
        now_str = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        log_subdir = os.path.join(log_dir, now_str)
        os.makedirs(log_subdir, exist_ok=True)
        log_filename = os.path.join(log_subdir, 'Log.txt')
        
        # 儲存時間戳和檔案路徑供後續使用
        LoggerManager._current_timestamp = now_str
        LoggerManager._current_log_file = log_filename
        
        file_handler = logging.FileHandler(log_filename, mode='w', encoding='utf-8')
        file_handler.setFormatter(file_formatter)
        logger.addHandler(file_handler)
        
        # 3. 主控台處理器 (可選，用於除錯)
        # console_handler = logging.StreamHandler()
        # console_handler.setFormatter(formatter)
        # logger.addHandler(console_handler)
        
        # 記錄初始訊息
        logger.info("=" * 60)
        logger.info("ICR 辨識率測試系統已啟動")
        logger.info("=" * 60)
        
        return logger
    
    @staticmethod
    def get_current_timestamp() -> Optional[str]:
        """取得當前的時間戳"""
        return LoggerManager._current_timestamp
    
    @staticmethod
    def get_current_log_file() -> Optional[str]:
        """取得當前的 log 檔案路徑"""
        return LoggerManager._current_log_file
    
    @staticmethod
    def get_logger() -> logging.Logger:
        """取得已配置的 Logger"""
        return logging.getLogger(LoggerManager.LOGGER_NAME)
    
    @staticmethod
    def log_section(title: str, logger: Optional[logging.Logger] = None):
        """記錄區段標題"""
        if logger is None:
            logger = LoggerManager.get_logger()
        
        logger.info("-" * 60)
        logger.info(f"【{title}】")
        logger.info("-" * 60)
    
    @staticmethod
    def log_step(step_num: int, total_steps: int, title: str, logger: Optional[logging.Logger] = None):
        """記錄步驟標題"""
        if logger is None:
            logger = LoggerManager.get_logger()
        
        logger.info(f"【步驟 {step_num}/{total_steps}】{title}")
        logger.info("-" * 60)

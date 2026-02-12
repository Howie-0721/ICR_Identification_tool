"""
資料庫匯出服務
提供資料表匯出至 CSV 的功能
"""

import os
import logging
import warnings
from typing import List, Tuple, Dict
import psycopg2
import pandas as pd


class DatabaseExporter:
    """資料庫匯出服務"""
    
    def __init__(self, host: str, port: any, database: str, user: str, password: str):
        """
        初始化資料庫匯出器
        
        Args:
            host: 資料庫主機位址
            port: 連接埠
            database: 資料庫名稱
            user: 使用者名稱
            password: 密碼
        """
        self.db_config = {
            'host': host,
            'port': port,
            'database': database,
            'user': user,
            'password': password
        }
        self.logger = logging.getLogger("ICRLogger")
    
    def export_tables(
        self,
        table_queries: List[Tuple[str, str]],
        output_dir: str
    ) -> bool:
        """
        匯出多個資料表到 CSV
        
        Args:
            table_queries: 表格名稱與查詢語句的列表
                          格式: [(table_name, query), ...]
            output_dir: 輸出目錄路徑
            
        Returns:
            成功返回 True，失敗時拋出異常
            
        Raises:
            Exception: 資料庫連接或查詢失敗時拋出
        """
        self.logger.info("開始匯出資料庫表格")
        
        # 忽略 pandas 警告
        warnings.filterwarnings(
            "ignore",
            message="pandas only supports SQLAlchemy connectable*",
            category=UserWarning
        )
        
        try:
            # 連接資料庫
            self.logger.debug(
                f"連接資料庫: {self.db_config['host']}:"
                f"{self.db_config['port']}/{self.db_config['database']}"
            )
            conn = psycopg2.connect(**self.db_config)
            self.logger.info("資料庫連接成功")
            
            # 建立輸出目錄
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # 逐一匯出表格
            for idx, (table_name, query) in enumerate(table_queries, 1):
                try:
                    self.logger.debug(f"查詢表格 [{idx}/{len(table_queries)}]: {table_name}")
                    
                    # 執行查詢
                    df = pd.read_sql(query, conn)
                    
                    # 寫入 CSV
                    csv_filename = os.path.join(output_dir, f"{table_name}.csv")
                    df.to_csv(csv_filename, index=False, encoding="utf-8-sig")
                    
                    self.logger.info(f"成功匯出 {table_name}: {len(df)} 筆記錄 → {csv_filename}")
                    
                except Exception as e:
                    self.logger.error(f"取得表 {table_name} 失敗: {e}")
            
            # 關閉連接
            conn.close()
            self.logger.info("資料庫匯出完成")
            return True
            
        except Exception as e:
            self.logger.error(f"資料庫連接失敗: {e}")
            raise

"""
資料庫資料取得腳本
使用 DatabaseExporter 從資料庫匯出 4 個表格的全資料到 CSV
"""

import os
import sys
import configparser
from datetime import datetime

# 加入專案根目錄到路徑
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, project_root)

from core.config import DatabaseConfig, PathConfig
from processors.database_service import DatabaseExporter


def main():
    """主函數：取得資料庫資料並匯出到 CSV"""
    
    # 取得工作目錄
    work_dir = os.getcwd()
    
    # 讀取配置檔案
    config_file = os.path.join(work_dir, 'config.ini')
    config_parser = configparser.ConfigParser()
    config_parser.read(config_file, encoding='utf-8')
    
    # 建立配置物件
    db_config_obj = DatabaseConfig()
    path_config = PathConfig(work_dir=work_dir)
    
    # 取得資料庫連線配置
    db_connection_config = db_config_obj.get_connection_config(config_parser)
    
    # 產生時間戳記
    timestamp = datetime.now().strftime("%Y_%m_%d_%H%M%S")
    
    # 取得表格查詢列表，並加入時間戳記到檔案名稱
    table_queries = [(f"{table_name}_{timestamp}", query) for table_name, query in db_config_obj.tables]
    
    # 建立資料庫匯出器
    exporter = DatabaseExporter(**db_connection_config)
    
    # 匯出資料到 DB 目錄
    output_dir = path_config.db_dir
    exporter.export_tables(table_queries, output_dir)
    
    print("資料匯出完成！")


if __name__ == "__main__":
    main()

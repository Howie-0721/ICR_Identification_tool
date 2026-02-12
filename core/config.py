"""
配置管理模組
負責所有系統配置的集中管理，包括 URL、路徑、資料庫配置和文件類型配置
"""

import os
import configparser
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple


# ============================================================================
# 文件類型配置
# ============================================================================

@dataclass
class DocumentTypeConfig:
    """文件類型標準化配置"""
    name: str
    upload_folder: str
    answer_file: str
    doc_csv: str
    fields: List[str]
    doc_type_value: str
    field_mapping: Dict[str, str]
    output_columns: Optional[List[str]] = None
    is_employment: bool = False


# ============================================================================
# URL 配置
# ============================================================================

@dataclass
class URLConfig:
    """URL 配置類別"""
    login: str = "https://192.168.160.113/systalk/login"
    systalk: str = "https://192.168.160.113/systalk/"
    icr: str = "http://192.168.160.113/icr"
    document_manage: str = "https://192.168.160.113/icr/document-manage"


# ============================================================================
# 路徑配置
# ============================================================================

@dataclass
class PathConfig:
    """路徑配置類別"""
    work_dir: str = field(default_factory=os.getcwd)
    db_dir: str = field(init=False)
    config_file: str = 'config.ini'
    default_result_filename: str = 'result.xlsx'
    log_dir: str = field(init=False)
    
    def __post_init__(self):
        """初始化後自動計算衍生路徑"""
        self.db_dir = os.path.join(self.work_dir, 'DB')
        self.log_dir = os.path.join(self.work_dir, 'Log')
    
    def get_answer_dir(self) -> str:
        """取得答案檔案目錄"""
        return os.path.join(self.work_dir, 'Answer')
    
    def get_upload_dir(self) -> str:
        """取得上傳檔案根目錄"""
        return os.path.join(self.work_dir, 'Upload_folder')
    
    def get_log_dir(self) -> str:
        """取得日誌目錄"""
        return self.log_dir
    
    def set_log_dir(self, log_dir: str) -> None:
        """設定日誌目錄"""
        self.log_dir = log_dir


# ============================================================================
# 資料庫配置
# ============================================================================

@dataclass
class DatabaseConfig:
    """資料庫配置類別"""
    tables: List[Tuple[str, str]] = field(default_factory=lambda: [
        ("doc_ARC", 'SELECT * FROM document."doc_ARC";'),
        ("doc_employment_approval", 'SELECT * FROM document."doc_employment_approval";'),
        ("doc_health_report", 'SELECT * FROM document."doc_health_report";'),
        ("document_master", 'SELECT * FROM document."document_master";'),
    ])
    
    def get_connection_config(self, config_parser: configparser.ConfigParser) -> Dict[str, any]:
        """從 ConfigParser 取得資料庫連線配置"""
        return {
            'host': config_parser['DATABASE']['host'],
            'port': config_parser['DATABASE']['port'],
            'database': config_parser['DATABASE']['database'],
            'user': config_parser['DATABASE']['user'],
            'password': config_parser['DATABASE']['password']
        }


# ============================================================================
# 文件類型標準配置 (TYPE_CONFIG)
# ============================================================================

def _create_type_config() -> Dict[str, DocumentTypeConfig]:
    """建立文件類型配置字典"""
    return {
        '1': DocumentTypeConfig(
            name='ARC',
            upload_folder=os.path.join('Upload_folder', 'ARC'),
            answer_file=os.path.join('Answer', 'ARC_Answer.xlsx'),
            doc_csv='doc_ARC.csv',
            fields=['資料類型', '居留效期', '居留證號', '核發日期', '舊式統一證號', '護照號碼', '雇主名稱'],
            doc_type_value='ARC',
            output_columns=['資料序號', '檔名', '資料類型', '居留證號', '核發日期', '居留效期', '舊式統一證號', '護照號碼', '雇主名稱'],
            field_mapping={
                '居留證號': 'field_arc_no',
                '核發日期': 'field_issue_date',
                '居留效期': 'field_expiry_date',
                '舊式統一證號': 'field_original_arc_no',
                '護照號碼': 'field_passport_no',
                '雇主名稱': 'field_employer_name'
            }
        ),
        '2': DocumentTypeConfig(
            name='Health',
            upload_folder=os.path.join('Upload_folder', 'Health'),
            answer_file=os.path.join('Answer', 'Health_Answer.xlsx'),
            doc_csv='doc_health_report.csv',
            fields=['文件類型', '體檢日期', '報告日期', '是否合格', '護照號碼', '雇主名稱'],
            doc_type_value='HEALTH_REPORT',
            output_columns=['資料序號', '檔名', '文件類型', '護照號碼', '體檢日期', '報告日期', '是否合格', '雇主名稱'],
            field_mapping={
                '護照號碼': 'field_passport_no',
                '體檢日期': 'field_examination_date',
                '報告日期': 'field_report_date',
                '是否合格': 'field_health_summary',
                '雇主名稱': 'field_employer_name'
            }
        ),
        '3': DocumentTypeConfig(
            name='Employment',
            upload_folder=os.path.join('Upload_folder', 'Employment'),
            answer_file=os.path.join('Answer', 'Employment_approval_Answer.xlsx'),
            doc_csv='doc_employment_approval.csv',
            fields=['文件類型', '聘可函號', '聘可發文日', '聘可收文日', '編號', '護照號碼', '工作起日', '工作迄日', '雇主名稱'],
            doc_type_value='EMPLOYMENT_APPROVAL',
            output_columns=['檔名', '文件類型', '雇主名稱', '聘可函號', '編號', '聘可發文日', '聘可收文日', '護照號碼', '工作起日', '工作迄日'],
            field_mapping={},
            is_employment=True
        )
    }


# ============================================================================
# 統一配置管理器 (單例模式)
# ============================================================================

class ConfigManager:
    """
    統一配置管理器 (Singleton Pattern)
    
    集中管理所有系統配置，包括：
    - URL 配置
    - 路徑配置
    - 資料庫配置
    - 文件類型配置
    - 執行時配置 (從 config.ini 載入)
    """
    
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._initialized = False
        return cls._instance
    
    def __init__(self):
        if self._initialized:
            return
        
        # 初始化基礎配置
        self.urls = URLConfig()
        self.paths = PathConfig()
        self.database = DatabaseConfig()
        self.type_config = _create_type_config()
        
        # 執行時配置 (從 config.ini 載入)
        self.runtime_config = None
        
        self._initialized = True
    
    def load_runtime_config(self, config_path: Optional[str] = None) -> configparser.ConfigParser:
        """
        從 config.ini 載入執行時配置
        
        Args:
            config_path: 配置檔案路徑，若為 None 則使用預設路徑
            
        Returns:
            ConfigParser 物件
        """
        if config_path is None:
            config_path = os.path.join(self.paths.work_dir, self.paths.config_file)
        
        config = configparser.ConfigParser()
        config.read(config_path, encoding='utf-8')
        self.runtime_config = config
        return config
    
    def get_sftp_config(self, override_config: Optional[Dict[str, str]] = None) -> Dict[str, str]:
        """取得 SFTP 配置"""
        if self.runtime_config is None:
            self.load_runtime_config()
        
        # 如果有覆蓋設定，優先使用覆蓋設定
        if override_config:
            return {
                'host': override_config.get('host', self.runtime_config['SFTP'].get('host', '')),
                'port': override_config.get('port', self.runtime_config['SFTP'].get('port', '')),
                'username': override_config.get('username', self.runtime_config['SFTP'].get('username', '')),
                'password': override_config.get('password', self.runtime_config['SFTP'].get('password', '')),
                'remote_path': override_config.get('remote_path', self.runtime_config['SFTP'].get('remote_path', ''))
            }
        
        return {
            'host': self.runtime_config['SFTP']['host'],
            'port': self.runtime_config['SFTP']['port'],
            'username': self.runtime_config['SFTP']['username'],
            'password': self.runtime_config['SFTP']['password'],
            'remote_path': self.runtime_config['SFTP'].get('remote_path')
        }
    
    def get_db_config(self, override_config: Optional[Dict[str, any]] = None) -> Dict[str, any]:
        """取得資料庫配置"""
        if self.runtime_config is None:
            self.load_runtime_config()
        
        # 如果有覆蓋設定，優先使用覆蓋設定
        if override_config:
            return {
                'host': override_config.get('host', self.runtime_config['DATABASE'].get('host', '')),
                'port': override_config.get('port', self.runtime_config['DATABASE'].get('port', '')),
                'database': override_config.get('database', self.runtime_config['DATABASE'].get('database', '')),
                'user': override_config.get('user', self.runtime_config['DATABASE'].get('user', '')),
                'password': override_config.get('password', self.runtime_config['DATABASE'].get('password', ''))
            }
        
        return self.database.get_connection_config(self.runtime_config)
    
    
    def get_doc_type_config(self, doc_type: str) -> DocumentTypeConfig:
        """取得指定文件類型的配置"""
        return self.type_config.get(doc_type)
    
    def set_log_dir(self, log_dir: str) -> None:
        """設定日誌目錄"""
        self.paths.set_log_dir(log_dir)


# ============================================================================
# 便利函式
# ============================================================================

def get_config_manager() -> ConfigManager:
    """取得 ConfigManager 單例實例"""
    return ConfigManager()

"""
測試流程協調器
負責串接所有業務模組，執行完整的測試流程
"""

import os
import json
import asyncio
import logging
import shutil
from typing import List, Dict, Optional, Callable, Tuple
from core.config import ConfigManager, DocumentTypeConfig
from processors import (
    SFTPUploader,
    RecognitionAutomation,
    DatabaseExporter,
    ExcelExporter,
    DataProcessor
)
from testing import FileValidator
from utils.file_helpers import read_excel_data, read_csv_data
from utils.data_helpers import ensure_list


class TestOrchestrator:
    """
    測試流程協調器
    
    負責完整測試流程的編排與執行：
    1. 準備答案檔案
    2. 準備待測檔案
    3. 檔名匹配驗證
    4. 上傳至 SFTP
    5. 執行辨識監控
    6. 匯出資料庫
    7. 資料處理與評分
    8. 匯出結果
    """
    
    def __init__(self, config_manager: Optional[ConfigManager] = None):
        """
        初始化協調器
        
        Args:
            config_manager: 配置管理器實例 (可選)
        """
        self.config = config_manager or ConfigManager()
        self.logger = logging.getLogger("ICRLogger")
        self.validator = FileValidator()  # 使用 testing 模組的驗證器
    
    def execute_test_workflow(
        self,
        doc_type: str,
        answer_file_path: str,
        upload_files: List[str],
        stop_check_callback: Optional[Callable[[], bool]] = None,
        sftp_config_override: Optional[Dict[str, str]] = None,
        db_config_override: Optional[Dict[str, any]] = None,
        answer_format: str = "分行呈現"
    ) -> Dict[str, any]:
        """
        執行完整測試流程
        
        Args:
            doc_type: 文件類型 ('1'=ARC, '2'=Health, '3'=Employment)
            answer_file_path: 答案檔案路徑
            upload_files: 待測檔案路徑列表
            stop_check_callback: 停止檢查回呼函式
            sftp_config_override: SFTP 設定覆蓋（優先使用）
            db_config_override: DB 設定覆蓋（優先使用）
            
        Returns:
            執行結果字典，包含 success, result_path, statistics
            
        Raises:
            Exception: 執行失敗時拋出
        """
        self.logger.info("=" * 60)
        self.logger.info("開始執行測試流程")
        self.logger.info("=" * 60)
        
        # 檢查是否需要停止
        if stop_check_callback and stop_check_callback():
            self.logger.warning("已取消執行")
            raise Exception("使用者取消執行")
        
        # 載入執行時配置
        self.config.load_runtime_config()
        config_entry = self.config.get_doc_type_config(doc_type)
        
        if not config_entry:
            raise Exception(f"無效的文件類型: {doc_type}")
        
        # Step 1: 準備答案檔案
        answer_data = self._prepare_answer_files(
            answer_file_path,
            config_entry,
            stop_check_callback
        )
        
        # Step 2: 準備待測檔案
        upload_full_path = self._prepare_upload_files(
            upload_files,
            config_entry,
            stop_check_callback
        )
        
        # Step 3: 檔名匹配驗證
        self._validate_file_matching(
            answer_data,
            upload_files,
            stop_check_callback
        )
        
        # Step 4: 上傳檔案並執行辨識
        upload_id = self._execute_upload_and_recognition(
            upload_full_path,
            stop_check_callback,
            sftp_config_override,
            db_config_override
        )
        
        # Step 5: 匯出資料庫
        self._export_database(upload_id, doc_type, stop_check_callback, db_config_override)
        
        # Step 6: 處理與評分，匯出結果
        result = self._process_and_export_results(
            doc_type,
            config_entry,
            answer_data,
            stop_check_callback,
            answer_format
        )
        
        self.logger.info("=" * 60)
        self.logger.info("測試流程執行完成")
        self.logger.info("=" * 60)
        
        return result
    
    def execute_no_answer_workflow(
        self,
        doc_type: str,
        upload_files: List[str],
        stop_check_callback: Optional[Callable[[], bool]] = None,
        sftp_config_override: Optional[Dict[str, str]] = None,
        db_config_override: Optional[Dict[str, any]] = None,
        answer_format: str = "分行呈現"
    ) -> Dict[str, any]:
        """
        執行無答案測試流程（只上傳和抓取DB，不對答案）
        
        Args:
            doc_type: 文件類型 ('1'=ARC, '2'=Health, '3'=Employment)
            upload_files: 待測檔案路徑列表
            stop_check_callback: 停止檢查回呼函式
            sftp_config_override: SFTP 設定覆蓋（優先使用）
            db_config_override: DB 設定覆蓋（優先使用）
            
        Returns:
            執行結果字典，包含 success, result_path
            
        Raises:
            Exception: 執行失敗時拋出
        """
        self.logger.info("=" * 60)
        self.logger.info("開始執行無答案測試流程")
        self.logger.info("=" * 60)
        
        # 檢查是否需要停止
        if stop_check_callback and stop_check_callback():
            self.logger.warning("已取消執行")
            raise Exception("使用者取消執行")
        
        # 載入執行時配置
        self.config.load_runtime_config()
        config_entry = self.config.get_doc_type_config(doc_type)
        
        if not config_entry:
            raise Exception(f"無效的文件類型: {doc_type}")
        
        # Step 1: 準備待測檔案
        upload_full_path = self._prepare_upload_files(
            upload_files,
            config_entry,
            stop_check_callback
        )
        
        # Step 2: 上傳檔案並執行辨識
        upload_id = self._execute_upload_and_recognition(
            upload_full_path,
            stop_check_callback,
            sftp_config_override,
            db_config_override
        )
        
        # Step 3: 匯出資料庫
        self._export_database(upload_id, doc_type, stop_check_callback, db_config_override)
        
        # Step 4: 處理並匯出結果（不對答案）
        result = self._process_and_export_no_answer_results(
            doc_type,
            config_entry,
            stop_check_callback,
            answer_format
        )
        
        self.logger.info("=" * 60)
        self.logger.info("無答案測試流程執行完成")
        self.logger.info("=" * 60)
        
        return result
    
    def _prepare_answer_files(
        self,
        answer_file_path: str,
        config_entry: DocumentTypeConfig,
        stop_check_callback: Optional[Callable] = None
    ) -> List[Dict]:
        """準備答案檔案"""
        from core.logger import LoggerManager
        LoggerManager.log_step(1, 5, "準備答案檔案")
        
        if stop_check_callback and stop_check_callback():
            raise Exception("使用者取消執行")
        
        # 建立答案目錄
        answer_dir = self.config.paths.get_answer_dir()
        os.makedirs(answer_dir, exist_ok=True)
        
        # 複製答案檔案
        target_answer_path = os.path.join(answer_dir, os.path.basename(config_entry.answer_file))
        shutil.copy2(answer_file_path, target_answer_path)
        self.logger.info(f"答案檔案已複製: {os.path.basename(config_entry.answer_file)}")
        
        # 讀取答案資料（不再重複記錄檔名，因為在 GUI 選擇時已記錄）
        answer_data = read_excel_data(target_answer_path)
        
        return answer_data
    
    def _prepare_upload_files(
        self,
        upload_files: List[str],
        config_entry: DocumentTypeConfig,
        stop_check_callback: Optional[Callable] = None
    ) -> str:
        """準備待測檔案"""
        from core.logger import LoggerManager
        LoggerManager.log_step(2, 5, "準備待測文件")
        
        if stop_check_callback and stop_check_callback():
            raise Exception("使用者取消執行")
        
        # 建立上傳目錄
        upload_full_path = os.path.join(self.config.paths.work_dir, config_entry.upload_folder)
        os.makedirs(upload_full_path, exist_ok=True)
        
        # 清空目錄
        for f in os.listdir(upload_full_path):
            file_to_remove = os.path.join(upload_full_path, f)
            if os.path.isfile(file_to_remove):
                os.remove(file_to_remove)
        
        # 複製待測檔案
        for file_path in upload_files:
            filename = os.path.basename(file_path)
            target_path = os.path.join(upload_full_path, filename)
            shutil.copy2(file_path, target_path)
            self.logger.info(f"已複製待測文件: {filename}")
        
        self.logger.info(f"共 {len(upload_files)} 個待測文件已準備完成")
        
        return upload_full_path
    
    def _validate_file_matching(
        self,
        answer_data: List[Dict],
        upload_files: List[str],
        stop_check_callback: Optional[Callable] = None
    ) -> None:
        """檔名匹配驗證（使用 testing.FileValidator）"""
        from core.logger import LoggerManager
        LoggerManager.log_section("檔名匹配驗證")
        
        if stop_check_callback and stop_check_callback():
            raise Exception("使用者取消執行")
        
        # 使用 FileValidator 進行驗證
        validation_result = self.validator.validate_file_matching(answer_data, upload_files)
        
        if not validation_result['valid']:
            # 產生錯誤訊息
            error_msg = self.validator.format_error_message(validation_result)
            
            self.logger.error("=" * 60)
            self.logger.error("檔名匹配驗證失敗，流程終止")
            self.logger.error("=" * 60)
            
            raise Exception(error_msg)
    
    def _execute_upload_and_recognition(
        self,
        upload_path: str,
        stop_check_callback: Optional[Callable] = None,
        sftp_config_override: Optional[Dict[str, str]] = None,
        db_config_override: Optional[Dict[str, any]] = None
    ) -> str:
        """
        上傳檔案並執行辨識
        
        Returns:
            upload_id: 上傳批次 ID
        """
        from core.logger import LoggerManager
        LoggerManager.log_step(3, 5, "上傳文件並執行辨識")
        
        if stop_check_callback and stop_check_callback():
            raise Exception("使用者取消執行")
        
        # 取得配置（優先使用覆蓋設定）
        sftp_config = self.config.get_sftp_config(sftp_config_override)
        db_config = self.config.get_db_config(db_config_override)
        # login_creds 已移除，不再需要
        
        # 上傳檔案
        try:
            if not sftp_config.get('remote_path'):
                raise Exception("SFTP remote_path 未設定於 config.ini")
            
            sftp_service = SFTPUploader(
                sftp_config['host'],
                int(sftp_config['port']),
                sftp_config['username'],
                sftp_config['password']
            )
            sftp_service.upload_folder(upload_path, sftp_config['remote_path'])
            self.logger.info("文件上傳完成")
            
        except Exception as e:
            self.logger.error(f"上傳文件失敗: {e}")
            raise
        
        if stop_check_callback and stop_check_callback():
            raise Exception("使用者取消執行")
        
        # 執行辨識
        try:
            web_automation = RecognitionAutomation(
                self.config.urls.login,
                self.config.urls.systalk,
                self.config.urls.icr,
                self.config.urls.document_manage,
                stop_check_callback
            )
            upload_id = asyncio.run(web_automation.monitor_and_recognize(
                db_config=db_config
            ))
            return upload_id
            
        except Exception as e:
            self.logger.error(f"辨識監控失敗: {e}")
            raise
    
    def _export_database(
        self,
        upload_id: str,
        doc_type: str,
        stop_check_callback: Optional[Callable] = None,
        db_config_override: Optional[Dict[str, any]] = None
    ) -> None:
        """匯出資料庫"""
        from core.logger import LoggerManager
        LoggerManager.log_step(4, 5, "匯出資料庫")
        
        if stop_check_callback and stop_check_callback():
            raise Exception("使用者取消執行")
        
        try:
            db_config = self.config.get_db_config(db_config_override)
            db_exporter = DatabaseExporter(
                db_config['host'],
                db_config['port'],
                db_config['database'],
                db_config['user'],
                db_config['password']
            )
            
            # 根據 upload_id 動態生成過濾查詢
            config_entry = self.config.get_doc_type_config(doc_type)
            self.logger.info(f"使用 Upload ID 過濾資料庫: {upload_id}")
            tables_with_filter = self._build_filtered_queries(upload_id, config_entry)
            
            db_exporter.export_tables(
                tables_with_filter,
                self.config.paths.db_dir
            )
            
        except Exception as e:
            self.logger.error(f"匯出資料庫失敗: {e}")
            raise
    
    def _build_filtered_queries(
        self,
        upload_id: str,
        config_entry: DocumentTypeConfig
    ) -> List[Tuple[str, str]]:
        """
        根據 upload_id 構建過濾後的資料庫查詢語句
        
        Args:
            upload_id: 上傳批次 ID
            config_entry: 文件類型配置
            
        Returns:
            包含表名和查詢語句的元組列表
        """
        # document_master 表過濾 (file_storage_path 包含 upload_id + COMPLETED 狀態)
        doc_master_query = f'''
            SELECT * FROM document."document_master"
            WHERE file_storage_path LIKE '%/{upload_id}/%'
            AND recognition_status = 'COMPLETED';
        '''
        
        # 文件類型特定表過濾 (透過 uuid 關聯)
        doc_type_query = f'''
            SELECT dt.* FROM document."{config_entry.doc_csv.replace('.csv', '')}" dt
            INNER JOIN document."document_master" dm
            ON dt.uuid = dm.uuid
            WHERE dm.file_storage_path LIKE '%/{upload_id}/%'
            AND dm.recognition_status = 'COMPLETED';
        '''
        
        # 其他文件類型表也需要過濾 (避免混入其他批次資料)
        other_doc_types = ['doc_ARC', 'doc_employment_approval', 'doc_health_report']
        current_doc_table = config_entry.doc_csv.replace('.csv', '')
        
        queries = [
            ("document_master", doc_master_query),
            (current_doc_table, doc_type_query)
        ]
        
        # 其他文件類型表匯出空結果 (或過濾當前 upload_id)
        for table in other_doc_types:
            if table != current_doc_table:
                empty_query = f'''
                    SELECT dt.* FROM document."{table}" dt
                    INNER JOIN document."document_master" dm
                    ON dt.uuid = dm.uuid
                    WHERE dm.file_storage_path LIKE '%/{upload_id}/%'
                    AND dm.recognition_status = 'COMPLETED';
                '''
                queries.append((table, empty_query))
        
        return queries
    
    def _process_and_export_results(
        self,
        doc_type: str,
        config_entry: DocumentTypeConfig,
        answer_data: List[Dict],
        stop_check_callback: Optional[Callable] = None,
        answer_format: str = "分行呈現"
    ) -> Dict[str, any]:
        """處理與評分，匯出結果"""
        from core.logger import LoggerManager
        LoggerManager.log_step(5, 5, "合併與評分")
        
        if stop_check_callback and stop_check_callback():
            raise Exception("使用者取消執行")
        
        # 處理與評分
        try:
            data_processor = DataProcessor(self.config.paths.db_dir)
            final_rows = data_processor.process_and_score(
                doc_type, 
                config_entry, 
                answer_data,
                answer_format=answer_format
            )
            
        except Exception as e:
            self.logger.error(f"處理與評分失敗: {e}")
            raise
        
        if final_rows is None:
            raise Exception("處理失敗，無法繼續")
        
        # 匯出結果
        LoggerManager.log_section("匯出結果")
        
        try:
            # 直接生成到 Log/timestamp/results.xlsx
            from core.logger import LoggerManager as LM
            timestamp = LM.get_current_timestamp()
            
            if not timestamp:
                raise Exception("無法取得時間戳，無法建立輸出目錄")
            
            log_dir = self.config.paths.get_log_dir()
            archive_dir = os.path.join(log_dir, timestamp)
            os.makedirs(archive_dir, exist_ok=True)
            output_path = os.path.join(archive_dir, 'results.xlsx')
            
            base_cols = config_entry.output_columns if config_entry.output_columns else []
            output_columns = DataProcessor.get_full_output_columns(final_rows, base_cols)
            
            excel_exporter = ExcelExporter()
            success = excel_exporter.export_to_excel(
                final_rows, 
                output_columns, 
                output_path,
                answer_data=answer_data,
                base_columns=base_cols,
                doc_type=config_entry.name
            )
            
            if not success:
                raise Exception("Excel 匯出失敗")
            
            # 統計結果
            stats = {
                'pass': sum(1 for row in final_rows if row.get('辨識結果') == 'PASS'),
                'fail': sum(1 for row in final_rows if row.get('辨識結果') == 'FAIL')
            }
            
            self.logger.info(f"評分結果: PASS {stats['pass']} / FAIL {stats['fail']}")
            
            # 歸檔測試記錄到時間戳資料夾
            self._archive_test_results(doc_type, config_entry, output_path)
            
            return {
                'success': True,
                'result_path': output_path,
                'statistics': stats
            }
            
        except Exception as e:
            self.logger.error(f"匯出結果失敗: {e}")
            raise
    
    def _process_and_export_no_answer_results(
        self,
        doc_type: str,
        config_entry: DocumentTypeConfig,
        stop_check_callback: Optional[Callable] = None,
        answer_format: str = "分行呈現"
    ) -> Dict[str, any]:
        """處理並匯出結果（不對答案）"""
        from core.logger import LoggerManager
        LoggerManager.log_section("匯出辨識結果（無答案）")
        
        if stop_check_callback and stop_check_callback():
            raise Exception("使用者取消執行")
        
        # 處理資料（不評分）
        try:
            # 讀取 document_master.csv
            document_master = read_csv_data(os.path.join(self.config.paths.db_dir, 'document_master.csv'))
            
            # 合併資料（根據文件類型使用不同邏輯）
            if config_entry.is_employment:
                output_rows = self._merge_employment_type_no_answer(document_master, answer_format)
            else:
                output_rows = self._merge_standard_type_no_answer(doc_type, config_entry, document_master)
            
            if not output_rows:
                raise Exception("沒有資料可匯出")
            
            # 按檔名排序
            output_rows = sorted(output_rows, key=lambda x: x.get('檔名', ''))
            
        except Exception as e:
            self.logger.error(f"處理資料失敗: {e}")
            raise
        
        if not output_rows:
            raise Exception("沒有資料可匯出")
        
        # 匯出結果
        LoggerManager.log_section("匯出結果到 Identification_results.xlsx")
        
        try:
            # 直接生成到 Log/timestamp/Identification_results.xlsx
            from core.logger import LoggerManager as LM
            timestamp = LM.get_current_timestamp()
            
            if not timestamp:
                raise Exception("無法取得時間戳，無法建立輸出目錄")
            
            log_dir = self.config.paths.get_log_dir()
            archive_dir = os.path.join(log_dir, timestamp)
            os.makedirs(archive_dir, exist_ok=True)
            output_path = os.path.join(archive_dir, 'Identification_results.xlsx')
            
            # 只使用基礎欄位（不包含答案欄位），無答案匯出不含資料序號
            base_cols = config_entry.output_columns if config_entry.output_columns else []
            output_columns = [col for col in base_cols if col != '資料序號']
            
            excel_exporter = ExcelExporter()
            success = excel_exporter.export_to_excel(
                output_rows, 
                output_columns, 
                output_path,
                answer_data=None,  # 不需要答案
                base_columns=base_cols,
                doc_type=config_entry.name
            )
            
            if not success:
                raise Exception("Excel 匯出失敗")
            
            self.logger.info(f"共匯出 {len(output_rows)} 筆辨識結果")
            
            # 歸檔測試記錄到時間戳資料夾
            self._archive_test_results(doc_type, config_entry, output_path)
            
            return {
                'success': True,
                'result_path': output_path
            }
            
        except Exception as e:
            self.logger.error(f"匯出結果失敗: {e}")
            raise
    
    def _archive_test_results(
        self,
        doc_type: str,
        config_entry: DocumentTypeConfig,
        result_path: str
    ) -> None:
        """
        將測試結果歸檔到時間戳資料夾
        
        建立結構：
        Log/20260122_110400/
            ├── Log.txt
            ├── DB/
            │   ├── doc_ARC.csv
            │   ├── doc_employment_approval.csv
            │   ├── doc_health_report.csv
            │   └── document_master.csv
            ├── test_data/（ARC 或 Health 或 Employment）
            │   ├── 文件1.pdf
            │   └── 文件2.pdf
            └── results.xlsx
        """
        from core.logger import LoggerManager
        
        try:
            # 取得時間戳
            timestamp = LoggerManager.get_current_timestamp()
            log_file = LoggerManager.get_current_log_file()
            
            if not timestamp or not log_file:
                self.logger.warning("無法取得時間戳資訊，跳過歸檔")
                return
            
            # 建立時間戳資料夾
            log_dir = self.config.paths.get_log_dir()
            archive_dir = os.path.join(log_dir, timestamp)
            os.makedirs(archive_dir, exist_ok=True)
            
            self.logger.info("=" * 60)
            self.logger.info(f"開始歸檔測試結果到: {timestamp}")
            self.logger.info("=" * 60)
            
            # 1. 確保 Log 檔案內容已 flush 並確認是否需要複製
            log_archive_path = os.path.join(archive_dir, 'Log.txt')
            logger = LoggerManager.get_logger()
            
            # Flush 所有 handler 確保資料寫入
            for handler in logger.handlers:
                if isinstance(handler, logging.FileHandler):
                    handler.flush()
            
            # 如果日誌檔案不在歸檔資料夾，才進行複製
            try:
                if os.path.normpath(log_file) != os.path.normpath(log_archive_path):
                    shutil.copy2(log_file, log_archive_path)
                    self.logger.info(f"已歸檔 Log.txt")
                else:
                    self.logger.info("Log.txt 已直接寫入於歸檔資料夾")
            except Exception as e:
                self.logger.warning(f"複製 Log 檔案到歸檔資料夾時發生問題: {e}")
            
            # 2. 複製 DB 資料夾
            db_source = self.config.paths.db_dir
            db_archive = os.path.join(archive_dir, 'DB')
            if os.path.exists(db_source):
                shutil.copytree(db_source, db_archive, dirs_exist_ok=True)
                csv_count = len([f for f in os.listdir(db_archive) if f.endswith('.csv')])
                self.logger.info(f"已歸檔 DB 資料夾 ({csv_count} 個 CSV 檔)")
            
            # 3. 複製 test_data 資料夾（根據文件類型，加入類型子資料夾）
            upload_source = os.path.join(self.config.paths.work_dir, config_entry.upload_folder)
            test_data_archive = os.path.join(archive_dir, 'test_data')
            if os.path.exists(upload_source):
                # 取得文件類型名稱 (ARC, Health, Employment)
                type_folder_name = os.path.basename(config_entry.upload_folder)  # 例如: "ARC", "Health", "Employment"
                
                # 建立 test_data/ARC (或 Health/Employment) 結構
                test_data_type_dir = os.path.join(test_data_archive, type_folder_name)
                os.makedirs(test_data_type_dir, exist_ok=True)
                
                # 複製文件到對應類型子資料夾
                for file_name in os.listdir(upload_source):
                    src_file = os.path.join(upload_source, file_name)
                    if os.path.isfile(src_file):
                        shutil.copy2(src_file, os.path.join(test_data_type_dir, file_name))
                
                file_count = len([f for f in os.listdir(test_data_type_dir) if os.path.isfile(os.path.join(test_data_type_dir, f))])
                self.logger.info(f"已歸檔 test_data/{type_folder_name} 資料夾 ({file_count} 個檔案)")
            
            # 4. 複製 results.xlsx（results.xlsx 已經在歸檔目錄中，不需要再複製）
            # output_path 就是 Log/timestamp/results.xlsx，所以不需要額外處理
            
            self.logger.info("=" * 60)
            self.logger.info(f"歸檔完成: {archive_dir}")
            self.logger.info("=" * 60)
            
        except Exception as e:
            self.logger.error(f"歸檔失敗: {e}")
            # 歸檔失敗不影響主流程，僅記錄錯誤
    
    def _merge_standard_type_no_answer(
        self,
        doc_type: str,
        config_entry: DocumentTypeConfig,
        document_master: List[Dict]
    ) -> List[Dict]:
        """標準類型（ARC、Health）的無答案合併邏輯"""
        # 讀取文件類型特定資料
        doc_csv_path = os.path.join(self.config.paths.db_dir, os.path.basename(config_entry.doc_csv))
        doc_data = read_csv_data(doc_csv_path)
        doc_dict = {row['uuid']: row for row in doc_data if 'uuid' in row}
        
        # 合併資料
        output_rows = []
        for row in document_master:
            uuid = row.get('uuid', '')
            output_row = {
                '檔名': row.get('file_name', '')
            }
            
            # 設定文件類型欄位
            if doc_type == '1':
                output_row['資料類型'] = row.get('document_type', '')
            else:
                output_row['文件類型'] = row.get('document_type', '')
            
            # 填入欄位資料
            if uuid in doc_dict:
                doc_row = doc_dict[uuid]
                for field, db_field in config_entry.field_mapping.items():
                    output_row[field] = doc_row.get(db_field, '')
            else:
                for field in config_entry.field_mapping.keys():
                    output_row[field] = ''
            
            output_rows.append(output_row)
        
        return output_rows
    
    def _merge_employment_type_no_answer(
        self,
        document_master: List[Dict],
        answer_format: str = "分行呈現"
    ) -> List[Dict]:
        """Employment 類型的無答案合併邏輯"""
        output_rows = []
        for row in document_master:
            llm_output = row.get('llm_output', '')
            if not llm_output:
                # 如果沒有 llm_output，創建空白記錄
                output_row = {
                    '檔名': row.get('file_name', ''),
                    '文件類型': row.get('document_type', ''),
                    '雇主名稱': '',
                    '聘可函號': '',
                    '編號': '',
                    '聘可發文日': '',
                    '聘可收文日': '',
                    '護照號碼': '',
                    '工作起日': '',
                    '工作迄日': ''
                }
                output_rows.append(output_row)
                continue
            
            try:
                data = json.loads(llm_output)
                
                numbers = ensure_list(data.get('編號', []))
                passports = ensure_list(data.get('護照號碼', []))
                start_dates = ensure_list(data.get('工作起日', []))
                end_dates = ensure_list(data.get('工作迄日', []))
                
                # 獲取基本資料
                file_name = row.get('file_name', '')
                approval_no = data.get('聘可函號', '')
                send_date = data.get('聘可發文日', '')
                receive_date = data.get('聘可收文日', '')
                employer_name = data.get('雇主名稱', '')
                
                # 根據答案形式決定輸出格式
                if answer_format == "列表呈現":
                    # 列表呈現：保持單行，使用 JSON 數組格式
                    output_row = {
                        '檔名': file_name,
                        '文件類型': row.get('document_type', ''),
                        '雇主名稱': employer_name,
                        '聘可函號': approval_no,
                        '編號': json.dumps(numbers, ensure_ascii=False) if numbers else '',
                        '聘可發文日': send_date,
                        '聘可收文日': receive_date,
                        '護照號碼': json.dumps(passports, ensure_ascii=False) if passports else '',
                        '工作起日': json.dumps(start_dates, ensure_ascii=False) if start_dates else '',
                        '工作迄日': json.dumps(end_dates, ensure_ascii=False) if end_dates else ''
                    }
                    output_rows.append(output_row)
                else:
                    # 分行呈現：展開列表為多列
                    max_len = max(len(numbers), len(passports), len(start_dates), len(end_dates))
                    
                    if max_len == 0:
                        # 沒有列表資料，創建一筆空白記錄
                        output_row = {
                            '檔名': file_name,
                            '文件類型': row.get('document_type', ''),
                            '雇主名稱': employer_name,
                            '聘可函號': approval_no,
                            '編號': '',
                            '聘可發文日': send_date,
                            '聘可收文日': receive_date,
                            '護照號碼': '',
                            '工作起日': '',
                            '工作迄日': ''
                        }
                        output_rows.append(output_row)
                    else:
                        # 展開為多列，每個編號一列
                        for i in range(max_len):
                            output_row = {
                                '檔名': file_name,
                                '文件類型': row.get('document_type', ''),
                                '雇主名稱': employer_name,
                                '聘可函號': approval_no,
                                '編號': numbers[i] if i < len(numbers) else '',
                                '聘可發文日': send_date,
                                '聘可收文日': receive_date,
                                '護照號碼': passports[i] if i < len(passports) else '',
                                '工作起日': start_dates[i] if i < len(start_dates) else '',
                                '工作迄日': end_dates[i] if i < len(end_dates) else ''
                            }
                            output_rows.append(output_row)
                
            except json.JSONDecodeError:
                # JSON 解析失敗，創建空白記錄
                output_row = {
                    '檔名': row.get('file_name', ''),                    '文件類型': row.get('document_type', ''),                    '雇主名稱': '',
                    '聘可函號': '',
                    '編號': '',
                    '聘可發文日': '',
                    '聘可收文日': '',
                    '護照號碼': '',
                    '工作起日': '',
                    '工作迄日': ''
                }
                output_rows.append(output_row)
        
        # 先按檔名排序，再按編號排序
        return sorted(output_rows, key=lambda x: (x.get('檔名', ''), x.get('編號', '')))

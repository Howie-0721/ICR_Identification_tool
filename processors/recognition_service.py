"""
辨識自動化服務
透過 API 呼叫與資料庫輪詢實現自動化辨識監控
"""

import asyncio
import logging
from typing import Dict, Optional, Callable
import aiohttp
import psycopg2


class RecognitionAutomation:
    """辨識自動化服務（使用 API + 資料庫輪詢）"""
    
    def __init__(
        self,
        login_url: str,
        systalk_url: str,
        icr_url: str,
        document_manage_url: str,
        stop_check_callback: Optional[Callable[[], bool]] = None
    ):
        """
        初始化辨識自動化服務
        
        Args:
            login_url: 登入 URL
            systalk_url: Systalk 系統 URL
            icr_url: ICR 系統 URL
            document_manage_url: 文件管理 URL
            stop_check_callback: 停止檢查回呼函式
        """
        self.login_url = login_url
        self.systalk_url = systalk_url
        self.icr_url = icr_url
        self.document_manage_url = document_manage_url
        self.logger = logging.getLogger("ICRLogger")
        self.api_url = "http://192.168.160.67:5003/api/v1/batchRecognition"
        self.regions = ["taipei"]
        self.stop_check_callback = stop_check_callback
    
    def get_completed_document_count(
        self,
        db_config: Dict[str, any],
        upload_id: Optional[str] = None
    ) -> int:
        """
        查詢資料庫中已完成辨識的文件數量 (recognition_status = 'COMPLETED')
        
        Args:
            db_config: 資料庫連線配置
            upload_id: 上傳批次 ID (可選)
            
        Returns:
            已完成文件數量，失敗時返回 -1
        """
        try:
            conn = psycopg2.connect(**db_config)
            cursor = conn.cursor()
            
            self.logger.debug(f"查詢參數 - upload_id: {upload_id}")
            
            if upload_id:
                # 使用更精確的路徑匹配，確保只匹配到該批次的檔案
                pattern = f'%/{upload_id}/%'
                self.logger.info(f"查詢 COMPLETED 狀態檔案 - 路徑模式: {pattern}")
                
                query = '''
                    SELECT COUNT(*), array_agg(file_storage_path) 
                    FROM document."document_master" 
                    WHERE recognition_status = 'COMPLETED' 
                    AND file_storage_path LIKE %s;
                '''
                cursor.execute(query, (pattern,))
                result = cursor.fetchone()
                count = result[0] if result[0] else 0
                paths = result[1] if result[1] else []
                
                # 記錄匹配到的檔案路徑（用於調試）
                if paths:
                    self.logger.debug(f"匹配到 {count} 個 COMPLETED 檔案:")
                    for path in paths[:5]:  # 只顯示前 5 個
                        self.logger.debug(f"  - {path}")
                    if len(paths) > 5:
                        self.logger.debug(f"  ... 還有 {len(paths) - 5} 個檔案")
                else:
                    self.logger.info(f"未找到符合條件的 COMPLETED 檔案（模式: {pattern}）")
            else:
                self.logger.warning("未提供 upload_id，將查詢所有 COMPLETED 記錄")
                query = '''
                    SELECT COUNT(*) 
                    FROM document."document_master" 
                    WHERE recognition_status = 'COMPLETED';
                '''
                cursor.execute(query)
                count = cursor.fetchone()[0]
            
            self.logger.info(f"查詢結果 - COMPLETED 數量: {count}")
            cursor.close()
            conn.close()
            return count
            
        except Exception as e:
            self.logger.error(f"查詢資料庫失敗: {e}")
            return -1
    
    async def call_batch_recognition_api(self) -> Optional[Dict[str, any]]:
        """
        呼叫整批辨識 API
        
        Returns:
            成功時返回 job 資訊字典，失敗時返回 None
        """
        self.logger.info(f"呼叫 API: {self.api_url}")
        self.logger.info(f"辨識區域: {self.regions}")
        
        try:
            async with aiohttp.ClientSession() as session:
                payload = {"regions": self.regions}
                async with session.post(self.api_url, json=payload) as response:
                    if response.status in [200, 201]:
                        result = await response.json()
                        self.logger.info(f"API 呼叫成功: {result}")
                        
                        if result.get('success'):
                            jobs = result.get('data', {}).get('jobs', [])
                            if jobs:
                                job = jobs[0]
                                return {
                                    'job_id': job.get('jobId'),
                                    'total_files': job.get('totalFiles'),
                                    'region': job.get('region'),
                                    'upload_id': job.get('uploadId')
                                }
                        return None
                    else:
                        error_text = await response.text()
                        self.logger.error(f"API 呼叫失敗: {response.status} - {error_text}")
                        return None
                        
        except Exception as e:
            self.logger.error(f"API 呼叫異常: {e}")
            return None
    
    async def poll_database_for_completion(
        self,
        db_config: Dict[str, any],
        initial_count: int,
        expected_increase: int,
        upload_id: Optional[str] = None,
        poll_interval: int = 20
    ) -> None:
        """
        輪詢資料庫直到辨識完成 (檢查 recognition_status = 'COMPLETED')
        
        Args:
            db_config: 資料庫連線配置
            initial_count: 初始已完成數量
            expected_increase: 預期新增數量
            upload_id: 上傳批次 ID (可選)
            poll_interval: 輪詢間隔 (秒)
            
        Raises:
            Exception: 使用者取消執行時拋出
        """
        self.logger.info(f"初始已完成文件數: {initial_count}, 預期新增: {expected_increase} 筆")
        if upload_id:
            self.logger.info(f"監控 Upload ID: {upload_id}")
        
        target_count = initial_count + expected_increase
        
        while True:
            # 檢查是否需要停止
            if self.stop_check_callback and self.stop_check_callback():
                self.logger.warning("偵測到終止請求，停止輪詢")
                raise Exception("使用者取消執行")
            
            # 查詢當前完成數量
            current_count = self.get_completed_document_count(db_config, upload_id)
            if current_count == -1:
                await asyncio.sleep(poll_interval)
                continue
            
            # 計算進度
            completed = current_count - initial_count
            progress = (completed / expected_increase * 100) if expected_increase > 0 else 0
            self.logger.info(f"辨識完成進度: {completed}/{expected_increase} ({progress:.1f}%)")
            
            # 檢查是否完成
            if current_count >= target_count:
                self.logger.info("所有檔案辨識完成！")
                break
            
            await asyncio.sleep(poll_interval)
    
    async def monitor_and_recognize(
        self,
        db_config: Optional[Dict[str, any]] = None
    ) -> None:
        """
        執行辨識（透過 API + 資料庫輪詢）
        
        Args:
            db_config: 資料庫配置 (可選)
            
        Raises:
            Exception: 辨識流程失敗時拋出
        """
        self.logger.info("開始執行自動化辨識操作")
        
        # 如果未提供 db_config，從 config.ini 載入
        if db_config is None:
            import configparser
            config = configparser.ConfigParser()
            config.read('config.ini', encoding='utf-8')
            db_config = {
                'host': config['DATABASE']['host'],
                'port': int(config['DATABASE']['port']),
                'database': config['DATABASE']['database'],
                'user': config['DATABASE']['user'],
                'password': config['DATABASE']['password']
            }
        
        try:
            # 步驟 1: 呼叫 API 啟動整批辨識
            self.logger.info("步驟 1/3: 呼叫 API 啟動整批辨識")
            job_info = await self.call_batch_recognition_api()
            if not job_info:
                raise Exception("API 呼叫失敗")
            
            upload_id = job_info.get('upload_id')
            
            # 驗證 upload_id
            if not upload_id:
                self.logger.error(f"API 返回的 Upload ID 為空: {job_info}")
                raise Exception("API 返回的 Upload ID 無效")
            
            self.logger.info(
                f"辨識任務已啟動 - JobID: {job_info['job_id']}, "
                f"Upload ID: {upload_id}, 總檔案數: {job_info['total_files']}"
            )
            
            # 步驟 2: 查詢資料庫初始狀態
            self.logger.info("步驟 2/3: 查詢資料庫初始狀態 (該批次已完成的文件數)")
            self.logger.info(f"查詢條件 - Upload ID: {upload_id}")
            initial_count = self.get_completed_document_count(db_config, upload_id)
            if initial_count == -1:
                raise Exception("無法連接資料庫")
            
            # 步驟 3: 輪詢資料庫監控辨識進度
            self.logger.info("步驟 3/3: 輪詢資料庫監控辨識進度")
            await self.poll_database_for_completion(
                db_config,
                initial_count,
                job_info['total_files'],
                upload_id,
                20
            )
            
            self.logger.info("辨識流程完成")
            return upload_id
            
        except Exception as e:
            self.logger.error(f"辨識流程失敗: {e}")
            raise

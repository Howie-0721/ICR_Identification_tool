"""
SFTP 上傳服務
提供檔案上傳至 SFTP 伺服器的功能
"""

import os
import logging
import paramiko
from typing import List


class SFTPUploader:
    """SFTP 檔案上傳服務"""
    
    def __init__(self, host: str, port: int, username: str, password: str):
        """
        初始化 SFTP 上傳器
        
        Args:
            host: SFTP 伺服器位址
            port: 連接埠
            username: 使用者名稱
            password: 密碼
        """
        self.host = host
        self.port = int(port)
        self.username = username
        self.password = password
        self.logger = logging.getLogger("ICRLogger")
    
    def upload_file(self, local_file_path: str, remote_path: str) -> None:
        """
        上傳單個文件到 SFTP 伺服器
        
        Args:
            local_file_path: 本地檔案路徑
            remote_path: 遠端目錄路徑
            
        Raises:
            Exception: 上傳失敗時拋出
        """
        self.logger.debug(f"開始上傳文件: {local_file_path}")
        
        try:
            # 建立 SSH 連接
            ssh = paramiko.SSHClient()
            ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            
            self.logger.debug(f"連接至 SFTP: {self.host}:{self.port}")
            ssh.connect(
                self.host,
                port=self.port,
                username=self.username,
                password=self.password
            )
            self.logger.info(f"已連接至 {self.host}:{self.port}")
            
            # 開啟 SFTP 連接
            sftp = ssh.open_sftp()
            remote_file_path = os.path.join(remote_path, os.path.basename(local_file_path))
            
            # 上傳檔案
            sftp.put(local_file_path, remote_file_path)
            self.logger.info(f"文件已上傳至 {remote_file_path}")
            
            # 關閉連接
            sftp.close()
            ssh.close()
            self.logger.debug("SFTP 連接已關閉")
            
        except Exception as e:
            self.logger.error(f"上傳失敗: {e}")
            raise
    
    def upload_folder(self, folder_path: str, remote_path: str) -> bool:
        """
        上傳資料夾內所有文件
        
        Args:
            folder_path: 本地資料夾路徑
            remote_path: 遠端目錄路徑
            
        Returns:
            上傳成功返回 True，否則返回 False
        """
        self.logger.info(f"開始上傳資料夾: {folder_path}")
        
        # 檢查資料夾是否存在
        if not os.path.exists(folder_path):
            self.logger.error(f"上傳資料夾不存在: {folder_path}")
            return False
        
        # 取得所有檔案
        files_to_upload = [
            f for f in os.listdir(folder_path)
            if os.path.isfile(os.path.join(folder_path, f))
        ]
        
        if not files_to_upload:
            self.logger.warning(f"資料夾中沒有文件: {folder_path}")
            return False
        
        self.logger.info(f"找到 {len(files_to_upload)} 個文件待上傳")
        
        # 逐一上傳
        for idx, file_name in enumerate(files_to_upload, 1):
            self.logger.info(f"上傳進度: [{idx}/{len(files_to_upload)}] {file_name}")
            local_file_path = os.path.join(folder_path, file_name)
            self.upload_file(local_file_path, remote_path)
        
        self.logger.info("文件上傳完成")
        return True

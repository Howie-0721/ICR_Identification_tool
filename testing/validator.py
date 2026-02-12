"""
檔案驗證器
負責檔名匹配驗證
"""

import os
import logging
from typing import List, Dict, Set


class FileValidator:
    """檔案驗證器 - 檔名匹配檢查"""
    
    def __init__(self):
        """初始化驗證器"""
        self.logger = logging.getLogger("ICRLogger")
    
    def validate_file_matching(
        self,
        answer_data: List[Dict],
        upload_files: List[str]
    ) -> Dict[str, any]:
        """
        驗證答案檔案與待測檔案的檔名是否匹配
        
        Args:
            answer_data: 答案資料列表
            upload_files: 待測檔案路徑列表
            
        Returns:
            驗證結果字典：
            {
                'valid': bool,
                'answer_filenames': set,
                'upload_filenames': set,
                'missing_in_upload': set,
                'missing_in_answer': set
            }
        """
        # 取得答案檔案中的檔名
        answer_filenames = set()
        if isinstance(answer_data, list) and answer_data:
            for row in answer_data:
                if isinstance(row, dict) and row:
                    first_value = str(next(iter(row.values())))
                    answer_filenames.add(first_value)
        
        # 取得上傳檔案的檔名
        upload_filenames = set(os.path.basename(f) for f in upload_files)
        
        # 檢查差異
        missing_in_upload = answer_filenames - upload_filenames
        missing_in_answer = upload_filenames - answer_filenames
        
        result = {
            'valid': len(missing_in_upload) == 0 and len(missing_in_answer) == 0,
            'answer_filenames': answer_filenames,
            'upload_filenames': upload_filenames,
            'missing_in_upload': missing_in_upload,
            'missing_in_answer': missing_in_answer
        }
        
        # 記錄驗證結果
        if result['valid']:
            self.logger.info(f"檔名匹配驗證通過 ({len(answer_filenames)} 個檔案)")
        else:
            self.logger.error("檔案不一致，請確認是否有上傳完整")
            
            if missing_in_upload:
                self.logger.error(f"答案檔案中有 {len(missing_in_upload)} 個檔案在待測文件中找不到")
                for fname in sorted(missing_in_upload):
                    self.logger.error(f"  答案檔案中找到但待測文件缺少: {fname}")
            
            if missing_in_answer:
                self.logger.error(f"待測文件中有 {len(missing_in_answer)} 個檔案在答案檔案中找不到")
                for fname in sorted(missing_in_answer):
                    self.logger.error(f"  待測文件中找到但答案檔案缺少: {fname}")
        
        return result
    
    def format_error_message(
        self,
        validation_result: Dict[str, any]
    ) -> str:
        """
        格式化錯誤訊息
        
        Args:
            validation_result: 驗證結果字典
            
        Returns:
            格式化的錯誤訊息字串
        """
        error_msg = "❌ 檔名匹配驗證失敗！\n\n"
        
        if validation_result['missing_in_upload']:
            missing = validation_result['missing_in_upload']
            error_msg += f"答案檔案中有 {len(missing)} 個檔案在待測文件中找不到：\n"
            for fname in sorted(missing):
                error_msg += f"  - {fname}\n"
        
        if validation_result['missing_in_answer']:
            missing = validation_result['missing_in_answer']
            error_msg += f"\n待測文件中有 {len(missing)} 個檔案在答案檔案中找不到：\n"
            for fname in sorted(missing):
                error_msg += f"  - {fname}\n"
        
        return error_msg

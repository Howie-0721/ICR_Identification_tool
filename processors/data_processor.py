"""
資料處理服務
負責資料合併與處理（評分邏輯已移至 testing/scorer.py）
"""

import os
import json
import logging
from collections import defaultdict
from typing import List, Dict, Optional
from core.config import DocumentTypeConfig
from utils.file_helpers import read_csv_data
from utils.data_helpers import parse_date_str
from processors.excel_service import ExcelExporter
from testing.scorer import TestScorer


class DataProcessor:
    """資料處理服務（專注於資料合併）"""
    
    def __init__(self, db_dir: str):
        """
        初始化資料處理器
        
        Args:
            db_dir: 資料庫匯出檔案目錄
        """
        self.db_dir = db_dir
        self.logger = logging.getLogger("ICRLogger")
        self.scorer = TestScorer()  # 使用 testing 模組的評分器
    
    @staticmethod
    def get_full_output_columns(
        final_rows: List[Dict[str, any]],
        base_columns: Optional[List[str]]
    ) -> List[str]:
        """
        自動產生完整欄位（含 _答案 欄位與辨識結果）
        
        Args:
            final_rows: 最終資料列表
            base_columns: 基礎欄位列表
            
        Returns:
            完整欄位列表
        """
        base_cols = base_columns if base_columns else []
        extra_cols = []
        
        if final_rows:
            for col in base_cols:
                ans_col = f"{col}_答案"
                if ans_col in final_rows[0]:
                    extra_cols.append(ans_col)
        
        result = base_cols + extra_cols
        
        if final_rows and '辨識結果' in final_rows[0] and '辨識結果' not in result:
            result.append('辨識結果')
        
        return result
    
    def process_and_score(
        self,
        choice: str,
        config_entry: DocumentTypeConfig,
        answer_data: List[Dict[str, str]],
        answer_format: str = "分行呈現"
    ) -> Optional[List[Dict[str, any]]]:
        """
        處理並評分數據
        
        Args:
            choice: 文件類型選擇 ('1'=ARC, '2'=Health, '3'=Employment)
            config_entry: 文件類型配置
            answer_data: 答案資料
            answer_format: 答案形式 ('分行呈現' 或 '列表呈現')
            
        Returns:
            評分後的資料列表，失敗時返回 None
        """
        self.logger.info(f"開始處理資料 - 類型: {config_entry.name}")
        
        # 讀取 document_master.csv
        self.logger.debug("讀取 document_master.csv")
        document_master = read_csv_data(os.path.join(self.db_dir, 'document_master.csv'))
        
        # 確保 doc_csv 路徑正確
        if config_entry.doc_csv:
            config_entry.doc_csv = os.path.join(self.db_dir, os.path.basename(config_entry.doc_csv))
        
        # 合併資料
        self.logger.info("合併資料...")
        if config_entry.is_employment:
            output_rows = self._merge_employment_type(document_master, answer_data, answer_format, config_entry)
        else:
            output_rows = self._merge_standard_type(choice, config_entry, document_master, answer_data)
        
        if output_rows is None:
            self.logger.error("合併資料失敗")
            return None
        
        self.logger.info(f"合併完成，共 {len(output_rows)} 筆資料")
        
        # 評分（使用 testing 模組的評分器）
        self.logger.info("開始評分...")
        scored_rows = self.scorer.score_data(
            output_rows,
            answer_data,
            config_entry.fields,
            config_entry.doc_type_value
        )
        
        return scored_rows
    
    def _merge_standard_type(
        self,
        choice: str,
        config_entry: DocumentTypeConfig,
        document_master: List[Dict],
        answer_data: List[Dict]
    ) -> Optional[List[Dict]]:
        """
        標準類型（ARC、Health）的合併邏輯
        
        Args:
            choice: 文件類型選擇
            config_entry: 文件類型配置
            document_master: document_master 資料
            answer_data: 答案資料
            
        Returns:
            合併後的資料列表，失敗時返回 None
        """
        self.logger.debug(f"開始合併資料 - 類型: {config_entry.name}")
        
        # 取得答案檔案中的檔名集合
        file_names_answer = set(row.get('檔名', '') for row in answer_data if row.get('檔名', ''))
        N = len(answer_data)
        self.logger.debug(f"答案檔案數量: {N}")
        
        # 取得最新 N 筆 document_master 記錄
        matching_master = sorted(
            document_master,
            key=lambda x: parse_date_str(x.get('created_at', '')),
            reverse=True
        )[:N]
        file_names_master = set(row.get('file_name', '') for row in matching_master if row.get('file_name', ''))
        
        # 檢查檔名一致性
        if file_names_master != file_names_answer:
            self.logger.error("檔案不一致，請確認是否有上傳完整")
            self.logger.debug(f"答案檔案: {file_names_answer}")
            self.logger.debug(f"資料庫檔案: {file_names_master}")
            
            # 列出缺少的檔案
            missing_in_master = file_names_answer - file_names_master
            missing_in_answer = file_names_master - file_names_answer
            
            if missing_in_master:
                self.logger.error(f"資料庫缺少 {len(missing_in_master)} 個檔案")
                for name in sorted(missing_in_master):
                    self.logger.error(f"  - {name}")
            
            if missing_in_answer:
                self.logger.error(f"答案檔案缺少 {len(missing_in_answer)} 個檔案")
                for name in sorted(missing_in_answer):
                    self.logger.error(f"  - {name}")
            
            return None
        
        # 讀取文件類型特定資料
        doc_data = read_csv_data(config_entry.doc_csv)
        doc_dict = {row['uuid']: row for row in doc_data if 'uuid' in row}
        
        # 合併資料
        output_rows = []
        for row in matching_master:
            uuid = row.get('uuid', '')
            output_row = {
                '資料序號': uuid,
                '檔名': row.get('file_name', '')
            }
            
            # 設定文件類型欄位
            if choice == '1':
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
        
        return sorted(output_rows, key=lambda x: x['檔名'])
    
    def _merge_employment_type(
        self,
        document_master: List[Dict],
        answer_data: List[Dict],
        answer_format: str = "分行呈現",
        config_entry: DocumentTypeConfig = None
    ) -> Optional[List[Dict]]:
        """
        Employment 類型的合併邏輯
        
        Args:
            document_master: document_master 資料
            answer_data: 答案資料
            answer_format: 答案形式 ('分行呈現' 或 '列表呈現')
            config_entry: 文件類型配置
            
        Returns:
            合併後的資料列表，失敗時返回 None
        """
        file_names_answer = set(row.get('檔名', '') for row in answer_data if row.get('檔名', ''))
        N = len(answer_data)
        
        # 取得最新 N 筆記錄
        matching_master = sorted(
            document_master,
            key=lambda x: parse_date_str(x.get('created_at', '')),
            reverse=True
        )[:N]
        file_names_master = set(row.get('file_name', '') for row in matching_master if row.get('file_name', ''))
        
        # 檢查檔名一致性
        if file_names_master != file_names_answer:
            self.logger.error("檔案不一致，請確認是否有上傳完整")
            
            # 列出缺少的檔案
            missing_in_master = file_names_answer - file_names_master
            missing_in_answer = file_names_master - file_names_answer
            
            if missing_in_master:
                self.logger.error(f"資料庫缺少 {len(missing_in_master)} 個檔案")
                for name in sorted(missing_in_master):
                    self.logger.error(f"  - {name}")
            
            if missing_in_answer:
                self.logger.error(f"答案檔案缺少 {len(missing_in_answer)} 個檔案")
                for name in sorted(missing_in_answer):
                    self.logger.error(f"  - {name}")
            
            return None
        
        # 建立檔名到文件類型的映射（從 document_master 中獲取）
        file_to_doc_type = {}
        for row in matching_master:
            file_name = row.get('file_name', '')
            doc_type = row.get('document_type', '')
            if file_name:
                file_to_doc_type[file_name] = doc_type
        
        # 獲取預期的文件類型值（用於答案）
        expected_doc_type = config_entry.doc_type_value if config_entry else ''
        
        # 建立答案字典（檔名對應答案資料）
        # 檢查答案數據是否已經分行（同一檔名有多行）
        answer_by_file = defaultdict(list)
        for row in answer_data:
            file_name = row.get('檔名', '')
            if file_name:
                answer_by_file[file_name].append(row)
        
        def ensure_list(val):
            """確保值為列表格式"""
            if isinstance(val, list):
                return val
            if isinstance(val, str):
                try:
                    parsed = json.loads(val)
                    if isinstance(parsed, list):
                        return parsed
                except Exception:
                    pass
                return [val] if val else []
            return []
        
        # 解析 Employment 資料
        output_rows = []
        for row in matching_master:
            llm_output = row.get('llm_output', '')
            if not llm_output:
                continue
            
            try:
                data = json.loads(llm_output)
                
                # 從 DB 取得列表資料
                numbers = ensure_list(data.get('編號', []))
                passports = ensure_list(data.get('護照號碼', []))
                start_dates = ensure_list(data.get('工作起日', []))
                end_dates = ensure_list(data.get('工作迄日', []))
                
                # 獲取基本資料
                file_name = row.get('file_name', '')
                doc_type = row.get('document_type', '')
                approval_no = data.get('聘可函號', '')
                send_date = data.get('聘可發文日', '')
                receive_date = data.get('聘可收文日', '')
                employer_name = data.get('雇主名稱', '')
                
                # 從答案取得資料
                answer_rows_for_file = answer_by_file.get(file_name, [])
                
                # 判斷答案格式：已分行 vs 列表格式
                if len(answer_rows_for_file) > 1:
                    # 答案已經分行：每行一個編號
                    answer_numbers = [row.get('編號', '') for row in answer_rows_for_file]
                    answer_passports = [row.get('護照號碼', '') for row in answer_rows_for_file]
                    answer_start_dates = [row.get('工作起日', '') for row in answer_rows_for_file]
                    answer_end_dates = [row.get('工作迄日', '') for row in answer_rows_for_file]
                    # 基本欄位從第一行取得
                    answer_doc_type = expected_doc_type  # 使用預期值
                    answer_approval_no = answer_rows_for_file[0].get('聘可函號', '')
                    answer_send_date = answer_rows_for_file[0].get('聘可發文日', '')
                    answer_receive_date = answer_rows_for_file[0].get('聘可收文日', '')
                    answer_employer_name = answer_rows_for_file[0].get('雇主名稱', '')
                elif len(answer_rows_for_file) == 1:
                    # 答案可能是列表格式：字段值是 JSON 數組
                    answer_row = answer_rows_for_file[0]
                    answer_numbers = ensure_list(answer_row.get('編號', []))
                    answer_passports = ensure_list(answer_row.get('護照號碼', []))
                    answer_start_dates = ensure_list(answer_row.get('工作起日', []))
                    answer_end_dates = ensure_list(answer_row.get('工作迄日', []))
                    answer_doc_type = expected_doc_type  # 使用預期值
                    answer_approval_no = answer_row.get('聘可函號', '')
                    answer_send_date = answer_row.get('聘可發文日', '')
                    answer_receive_date = answer_row.get('聘可收文日', '')
                    answer_employer_name = answer_row.get('雇主名稱', '')
                else:
                    # 沒有答案資料
                    answer_numbers = []
                    answer_passports = []
                    answer_start_dates = []
                    answer_end_dates = []
                    answer_doc_type = expected_doc_type  # 使用預期值
                    answer_approval_no = ''
                    answer_send_date = ''
                    answer_receive_date = ''
                    answer_employer_name = ''
                
                # 檢查 DB 數據和答案數據的列表長度是否一致
                db_max_len = max(len(numbers), len(passports), len(start_dates), len(end_dates))
                answer_max_len = max(len(answer_numbers), len(answer_passports), len(answer_start_dates), len(answer_end_dates))
                
                if db_max_len > 0 and answer_max_len > 0 and db_max_len != answer_max_len:
                    self.logger.warning(f"檔案 {file_name} 的 DB 數據列表長度({db_max_len})與答案數據列表長度({answer_max_len})不一致")
                
                # 根據答案形式決定輸出格式
                if answer_format == "列表呈現":
                    # 列表呈現：保持單行，使用 JSON 數組格式
                    output_row = {
                        '檔名': file_name,
                        '文件類型': doc_type,
                        '雇主名稱': employer_name,
                        '聘可函號': approval_no,
                        '編號': json.dumps(numbers, ensure_ascii=False) if numbers else '',
                        '聘可發文日': send_date,
                        '聘可收文日': receive_date,
                        '護照號碼': json.dumps(passports, ensure_ascii=False) if passports else '',
                        '工作起日': json.dumps(start_dates, ensure_ascii=False) if start_dates else '',
                        '工作迄日': json.dumps(end_dates, ensure_ascii=False) if end_dates else '',
                        # 答案欄位（也是列表格式）
                        '文件類型_答案': answer_doc_type,
                        '雇主名稱_答案': answer_employer_name,
                        '聘可函號_答案': answer_approval_no,
                        '編號_答案': json.dumps(answer_numbers, ensure_ascii=False) if answer_numbers else '',
                        '聘可發文日_答案': answer_send_date,
                        '聘可收文日_答案': answer_receive_date,
                        '護照號碼_答案': json.dumps(answer_passports, ensure_ascii=False) if answer_passports else '',
                        '工作起日_答案': json.dumps(answer_start_dates, ensure_ascii=False) if answer_start_dates else '',
                        '工作迄日_答案': json.dumps(answer_end_dates, ensure_ascii=False) if answer_end_dates else ''
                    }
                    output_rows.append(output_row)
                else:
                    # 分行呈現：展開列表為多列
                    max_len = max(db_max_len, answer_max_len)
                    
                    if max_len == 0:
                        # 沒有列表資料，創建一筆空白記錄
                        output_row = {
                            '檔名': file_name,
                            '文件類型': doc_type,
                            '雇主名稱': employer_name,
                            '聘可函號': approval_no,
                            '編號': '',
                            '聘可發文日': send_date,
                            '聘可收文日': receive_date,
                            '護照號碼': '',
                            '工作起日': '',
                            '工作迄日': '',
                            # 答案欄位
                            '文件類型_答案': answer_doc_type,
                            '雇主名稱_答案': answer_employer_name,
                            '聘可函號_答案': answer_approval_no,
                            '編號_答案': '',
                            '聘可發文日_答案': answer_send_date,
                            '聘可收文日_答案': answer_receive_date,
                            '護照號碼_答案': '',
                            '工作起日_答案': '',
                            '工作迄日_答案': ''
                        }
                        output_rows.append(output_row)
                    else:
                        # 展開為多列，每個編號一列
                        for i in range(max_len):
                            output_row = {
                                '檔名': file_name,
                                '文件類型': doc_type,
                                '雇主名稱': employer_name,
                                '聘可函號': approval_no,
                                '編號': numbers[i] if i < len(numbers) else '',
                                '聘可發文日': send_date,
                                '聘可收文日': receive_date,
                                '護照號碼': passports[i] if i < len(passports) else '',
                                '工作起日': start_dates[i] if i < len(start_dates) else '',
                                '工作迄日': end_dates[i] if i < len(end_dates) else '',
                                # 答案欄位（展開為多列）
                                '文件類型_答案': answer_doc_type,
                                '雇主名稱_答案': answer_employer_name,
                                '聘可函號_答案': answer_approval_no,
                                '編號_答案': answer_numbers[i] if i < len(answer_numbers) else '',
                                '聘可發文日_答案': answer_send_date,
                                '聘可收文日_答案': answer_receive_date,
                                '護照號碼_答案': answer_passports[i] if i < len(answer_passports) else '',
                                '工作起日_答案': answer_start_dates[i] if i < len(answer_start_dates) else '',
                                '工作迄日_答案': answer_end_dates[i] if i < len(answer_end_dates) else ''
                            }
                            output_rows.append(output_row)
                
            except json.JSONDecodeError:
                continue
        
        # 先按檔名排序，再按編號排序
        return sorted(output_rows, key=lambda x: (x.get('檔名', ''), x.get('編號', '')))
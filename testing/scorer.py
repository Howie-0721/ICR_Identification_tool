"""
評分器
負責測試結果的評分與 PASS/FAIL 判定
"""

import logging
from typing import List, Dict


class TestScorer:
    """測試評分器 - 評分與結果判定"""
    
    def __init__(self):
        """初始化評分器"""
        self.logger = logging.getLogger("ICRLogger")
    
    def score_data(
        self,
        output_rows: List[Dict],
        answer_data: List[Dict],
        fields: List[str],
        doc_type_value: str
    ) -> List[Dict]:
        """
        對測試資料進行評分
        
        Args:
            output_rows: 待評分資料
            answer_data: 答案資料
            fields: 評分欄位列表
            doc_type_value: 文件類型值
            
        Returns:
            評分後的資料列表（每筆資料包含 PASS/FAIL 欄位）
        """
        answer_dict = {row['檔名']: row for row in answer_data if '檔名' in row}
        type_field = '資料類型' if '資料類型' in fields else '文件類型'
        
        updated_rows = []
        for row in output_rows:
            file_name = row.get('檔名', '')
            
            # 檢查是否已經包含答案欄位（Employment 類型已展開）
            has_embedded_answers = any(f'{field}_答案' in row for field in fields)
            
            # 如果沒有內嵌答案且找不到對應的答案資料，跳過
            if not has_embedded_answers and file_name not in answer_dict:
                continue
            
            answer_row = answer_dict.get(file_name, {}) if not has_embedded_answers else {}
            overall_pass = True
            
            for field in fields:
                raw_value = row.get(field, '')
                raw_str = str(raw_value).strip() if raw_value is not None else ''
                
                # 如果已經有內嵌答案，使用內嵌答案；否則從 answer_dict 查找
                if has_embedded_answers and f'{field}_答案' in row:
                    answer_value = str(row.get(f'{field}_答案', '') or '').strip()
                else:
                    answer_value = str(answer_row.get(field) or '').strip()
                
                # 兩者都為空
                if raw_str == '' and answer_value == '':
                    row[field] = 'N/A'
                    row[f'{field}_答案'] = 'PASS'
                    continue
                
                # 文件類型欄位特殊處理：直接與預期的 doc_type_value 比較
                if field == type_field:
                    if raw_str == doc_type_value:
                        row[f'{field}_答案'] = 'PASS'
                    else:
                        row[f'{field}_答案'] = 'FAIL'
                        overall_pass = False
                    continue
                
                # 實際值為空但答案不為空
                if raw_str == '' and answer_value != '':
                    row[field] = f"N/A({answer_value})"
                    row[f'{field}_答案'] = 'FAIL'
                    overall_pass = False
                # 值匹配
                elif raw_str == answer_value:
                    row[f'{field}_答案'] = 'PASS'
                # 值不匹配
                else:
                    row[f'{field}_答案'] = 'FAIL'
                    if field != type_field:
                        row[field] = f"{raw_str}({answer_value})"
                    overall_pass = False
            
            row['辨識結果'] = 'PASS' if overall_pass else 'FAIL'
            updated_rows.append(row)
        
        # 統計評分結果
        pass_count = sum(1 for r in updated_rows if r.get('辨識結果') == 'PASS')
        fail_count = sum(1 for r in updated_rows if r.get('辨識結果') == 'FAIL')
        self.logger.info(f"評分完成：PASS={pass_count}，FAIL={fail_count}")
        
        return updated_rows
    
    def calculate_statistics(
        self,
        scored_rows: List[Dict]
    ) -> Dict[str, int]:
        """
        計算測試統計
        
        Args:
            scored_rows: 評分後的資料
            
        Returns:
            統計字典 {'pass': int, 'fail': int, 'total': int}
        """
        pass_count = sum(1 for row in scored_rows if row.get('辨識結果') == 'PASS')
        fail_count = sum(1 for row in scored_rows if row.get('辨識結果') == 'FAIL')
        
        return {
            'pass': pass_count,
            'fail': fail_count,
            'total': pass_count + fail_count
        }

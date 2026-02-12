"""
答案比對器
負責辨識結果與標準答案的比對邏輯
"""

import logging
from typing import Any, Dict, List, Optional


class AnswerComparator:
    """答案比對器 - 比對邏輯封裝"""
    
    def __init__(self):
        """初始化比對器"""
        self.logger = logging.getLogger("ICRLogger")
    
    def compare_field(
        self,
        actual_value: Any,
        expected_value: Any,
        field_name: str,
        is_type_field: bool = False,
        doc_type_value: str = None
    ) -> Dict[str, Any]:
        """
        比對單一欄位值
        
        Args:
            actual_value: 實際值
            expected_value: 預期值（答案）
            field_name: 欄位名稱
            is_type_field: 是否為文件類型欄位
            doc_type_value: 文件類型標準值
            
        Returns:
            比對結果字典：
            {
                'match': bool,
                'display_value': str,
                'result': 'PASS' | 'FAIL'
            }
        """
        actual_str = str(actual_value).strip() if actual_value is not None else ''
        expected_str = str(expected_value or '').strip()
        
        # 情況 1: 兩者都為空
        if actual_str == '' and expected_str == '':
            return {
                'match': True,
                'display_value': 'N/A',
                'result': 'PASS'
            }
        
        # 情況 2: 文件類型欄位特殊處理
        if is_type_field and actual_str == doc_type_value:
            return {
                'match': True,
                'display_value': actual_str,
                'result': 'PASS'
            }
        
        # 情況 3: 實際值為空但答案不為空
        if actual_str == '' and expected_str != '':
            return {
                'match': False,
                'display_value': f"N/A({expected_str})",
                'result': 'FAIL'
            }
        
        # 情況 4: 值匹配
        if actual_str == expected_str:
            return {
                'match': True,
                'display_value': actual_str,
                'result': 'PASS'
            }
        
        # 情況 5: 值不匹配
        display_value = f"{actual_str}({expected_str})" if not is_type_field else actual_str
        return {
            'match': False,
            'display_value': display_value,
            'result': 'FAIL'
        }
    
    def compare_row(
        self,
        actual_row: Dict[str, Any],
        expected_row: Dict[str, Any],
        fields: List[str],
        doc_type_value: str
    ) -> Dict[str, Any]:
        """
        比對整列資料
        
        Args:
            actual_row: 實際資料列
            expected_row: 答案資料列
            fields: 要比對的欄位列表
            doc_type_value: 文件類型標準值
            
        Returns:
            比對結果字典，包含各欄位比對結果
        """
        type_field = '資料類型' if '資料類型' in fields else '文件類型'
        result = {'overall_pass': True}
        
        for field in fields:
            is_type_field = (field == type_field)
            comparison = self.compare_field(
                actual_row.get(field, ''),
                expected_row.get(field, ''),
                field,
                is_type_field,
                doc_type_value
            )
            
            result[field] = comparison['display_value']
            result[f'{field}_答案'] = comparison['result']
            
            if not comparison['match']:
                result['overall_pass'] = False
        
        result['辨識結果'] = 'PASS' if result['overall_pass'] else 'FAIL'
        return result

"""
資料處理輔助函式
提供日期解析和資料類型轉換功能
"""

import datetime
import json
from typing import Any, List


def parse_date_str(date_str: str) -> datetime.datetime:
    """
    解析日期字串為 datetime 物件
    
    Args:
        date_str: 日期字串 (ISO 格式)
        
    Returns:
        datetime 物件，若解析失敗則返回 datetime.min
    """
    try:
        return datetime.datetime.fromisoformat(date_str)
    except (ValueError, TypeError):
        return datetime.datetime.min


def ensure_list(val: Any) -> List:
    """
    確保值為列表類型
    
    Args:
        val: 任意類型的值
        
    Returns:
        列表類型的值
        
    Examples:
        >>> ensure_list([1, 2, 3])
        [1, 2, 3]
        >>> ensure_list('["a", "b"]')
        ['a', 'b']
        >>> ensure_list('single')
        ['single']
    """
    # 已經是列表
    if isinstance(val, list):
        return val
    
    # 嘗試解析 JSON 字串
    if isinstance(val, str):
        try:
            parsed = json.loads(val)
            if isinstance(parsed, list):
                return parsed
        except Exception:
            pass
        # 將字串包裝為單元素列表
        return [val]
    
    # 其他類型返回空列表
    return []

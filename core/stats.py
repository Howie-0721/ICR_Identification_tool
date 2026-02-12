"""
統計與樣式配置模組
定義 UI 樣式常量和統計資料結構
"""

from dataclasses import dataclass
from typing import Dict, Any


class UIStyles:
    """UI 樣式配置常量"""
    
    # 主題色彩
    PRIMARY_BLUE = '#4472C4'
    BUTTON_BLUE = '#5B9BD5'
    BUTTON_GREEN = '#70AD47'
    BUTTON_ORANGE = '#ED7D31'
    BUTTON_RED = '#C95E5E'
    
    # 背景與文字色
    BG_LIGHT_GRAY = '#f5f5f5'
    WHITE = 'white'
    GRAY = 'gray'
    GREEN = 'green'
    RED = 'red'
    
    # 字型設定
    FONT_TITLE = ('Arial', 20, 'bold')
    FONT_SECTION = ('Arial', 12, 'bold')
    FONT_NORMAL = ('Arial', 11)
    FONT_SMALL = ('Arial', 10)
    FONT_LOG = ('Consolas', 10)


@dataclass
class TestStatistics:
    """測試統計資料"""
    total: int = 0
    passed: int = 0
    failed: int = 0
    
    @property
    def pass_rate(self) -> float:
        """通過率 (0-100)"""
        if self.total == 0:
            return 0.0
        return (self.passed / self.total) * 100
    
    def to_dict(self) -> Dict[str, Any]:
        """轉換為字典"""
        return {
            'total': self.total,
            'pass': self.passed,
            'fail': self.failed,
            'pass_rate': round(self.pass_rate, 2)
        }
    
    def __str__(self) -> str:
        return f"Total: {self.total}, Pass: {self.passed}, Fail: {self.failed} ({self.pass_rate:.1f}%)"

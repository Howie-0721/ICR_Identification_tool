"""
ICR 辨識率測試系統 - 核心模組
Core modules for ICR recognition testing system
"""

__version__ = '2.0.0'
__author__ = 'ICR Team'

from .config import ConfigManager
from .logger import LoggerManager
from .stats import UIStyles

__all__ = ['ConfigManager', 'LoggerManager', 'UIStyles']

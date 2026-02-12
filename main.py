"""
ICR 辨識率測試工具
"""

import sys
import tkinter as tk
from gui.app import ICRModernApp

def main():
    try:
        app = ICRModernApp()
        app.mainloop()
    except Exception as e:
        print(f"應用程式啟動失敗: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()

# ICR Performance Test 打包與執行指南

## 🚀 快速開始

### 打包成 EXE

**最簡單的方式：**
```bash
# 雙擊執行
build.bat
```

**或使用命令列：**
```powershell
# 1. 安裝依賴
pip install -r requirements.txt

# 2. 執行打包
pyinstaller perf_test.spec --clean
```

### 打包完成後的檔案結構

```
dist\perf_test\
├── perf_test.exe          ← 主程式（雙擊執行）
├── config.ini             ← 配置檔案
├── README.txt             ← 使用說明
├── Answer\                ← 答案檔案目錄
├── DB\                    ← 資料庫匯出目錄
├── Log\                   ← 日誌與測試記錄
├── excel_template\        ← Excel 模板檔案
│   ├── ARC_Sample.xlsx
│   ├── Health_Sample.xlsx
│   └── Employment_Sample.xlsx
├── Upload_folder\         ← 上傳檔案目錄
│   ├── ARC\
│   ├── Health\
│   └── Employment\
└── _internal\             ← 程式依賴檔案（PyInstaller 自動生成）
```

## 📦 部署到其他電腦

1. 將整個 `dist\perf_test\` 資料夾複製到目標電腦
2. 確認 `config.ini` 設定正確
3. 雙擊 `perf_test.exe` 啟動

**✅ 優點：**
- 不需要安裝 Python
- 不需要安裝任何依賴套件
- 開箱即用

## 🛠️ 打包工具說明

| 檔案 | 用途 |
|------|------|
| `build.bat` | 自動化打包批次檔（推薦使用） |
| `build_and_test.bat` | 打包並測試 |
| `perf_test.spec` | PyInstaller 配置檔 |
| `requirements.txt` | Python 依賴套件清單 |
| `BUILD_INSTRUCTIONS.md` | 詳細打包說明 |

## ⚙️ 自訂打包選項

編輯 `perf_test.spec` 檔案：

```python
# 更改程式名稱
name='perf_test',  # 改成你想要的名稱

# 隱藏控制台視窗
console=False,  # True=顯示, False=隱藏

# 自訂圖示
icon='your_icon.ico',  # 指定 .ico 檔案路徑
```

## 🐛 疑難排解

### 打包失敗

```bash
# 清除舊檔案重新打包
rmdir /s /q build dist
build.bat
```

### 執行時錯誤

1. 檢查 `config.ini` 是否存在
2. 查看 `Log` 資料夾中的日誌
3. 確認網路連線正常

### 防毒軟體誤報

PyInstaller 打包的程式可能被防毒軟體標記，請：
- 將程式加入白名單
- 或暫時停用防毒軟體進行測試

## 📝 開發模式執行

如果要在開發環境中執行（不打包）：

```bash
# 安裝依賴
pip install -r requirements.txt

# 直接執行
python main.py
```

## 🔄 更新版本

修改程式碼後：
1. 執行 `build.bat` 重新打包
2. 新的 exe 會在 `dist\perf_test\` 目錄

## 📋 系統需求

**開發/打包環境：**
- Windows 10/11
- Python 3.8+
- 網路連線

**執行環境（打包後）：**
- Windows 10/11
- 網路連線至 SFTP 和資料庫伺服器
- 不需要安裝 Python

## 📞 支援

如有問題請查看 `BUILD_INSTRUCTIONS.md` 獲取更多資訊。

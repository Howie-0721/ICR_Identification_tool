# ICR Identification Tool（辨識/效能測試工具）

此專案為 Windows 桌面工具（Tkinter GUI），用來執行 ICR 辨識相關的測試流程、彙整結果並輸出報表。

## 快速開始（開發模式）

1. 安裝依賴

```powershell
pip install -r requirements.txt
```

2. 建立本機設定檔（重要：`config.ini` 可能包含帳密，請勿提交）

```powershell
copy .\config.example.ini .\config.ini
```

3. 執行

```powershell
python main.py
```

## 快速開始（打包 EXE）

（可選）若需要在沒有 Python 的電腦上執行，可用 PyInstaller 打包。

```powershell
# 最簡單：直接雙擊 build.bat
# 或命令列執行：
pyinstaller perf_test.spec --clean
```

## 設定檔說明

- `config.ini`：實際執行用設定（已被 `.gitignore` 忽略，避免把 SFTP/DB 帳密推到 GitHub）
- `config.example.ini`：範例設定（可提交），請複製後再改成你自己的 `config.ini`

常見需要調整：
- SFTP / DATABASE 連線資訊
- `log_path`（建議用相對路徑，例如 `./Log`）
- 測試檔案路徑（`answer_file` / `upload_files`）

## 專案資料夾用途

- `core/`：核心流程與共用能力（設定、日誌、流程 orchestrator、統計等）
- `gui/`：GUI 介面（主視窗與元件）
- `processors/`：各種服務/處理器（SFTP、DB、Excel、辨識服務、資料處理等）
- `utils/`：通用工具函式（檔案/資料處理輔助）
- `testing/`：測試結果比對、計分、驗證等模組

- `DB/`：工具執行需要的 CSV 資料（例如 `document_master.csv`、`doc_*.csv`）
- `excel_template/`：輸出/測試用的 Excel 模板

- `Upload_folder/`：放待測的上傳檔案（依類型分資料夾：`ARC/`、`Health/`、`Employment/`）
- `Answer/`：放答案檔/比對用檔案（此資料夾內容通常不建議上版控）
- `Log/`：每次執行產生的輸出與日誌（時間戳記資料夾）

- `build/`、`dist/`：PyInstaller 產出物（不建議上版控）

## 常見輸出位置

- 執行記錄/結果：通常會在 `Log/<timestamp>/` 下面生成（依設定與流程而定）

## 版本控管注意事項

- 不要提交：`config.ini`（含帳密/內網位址）、`Log/`、`dist/`、`build/`、上傳檔與個人答案檔
- 已透過 `.gitignore` 做預設保護；若你新增了新的輸出資料夾，建議一併補到 `.gitignore`

## 問題排查

- 啟動失敗：先確認已建立 `config.ini`，並檢查 `Log/` 內最新一次執行的輸出
- 找不到資料：確認 `DB/` 內 CSV 是否齊全、`Upload_folder/` 路徑是否正確

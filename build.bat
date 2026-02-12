@echo off
chcp 65001 >nul
echo ========================================
echo ICR Performance Test 打包工具
echo ========================================
echo.

REM 檢查是否已安裝 PyInstaller
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo [錯誤] PyInstaller 未安裝
    echo 正在安裝依賴套件...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo [錯誤] 安裝失敗，請手動執行: pip install -r requirements.txt
        pause
        exit /b 1
    )
)

echo [步驟 1/4] 清理舊的打包檔案...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
echo 完成！
echo.

echo [步驟 2/4] 開始打包應用程式...
pyinstaller perf_test.spec --clean
if errorlevel 1 (
    echo [錯誤] 打包失敗！
    pause
    exit /b 1
)
echo 完成！
echo.

echo [步驟 3/4] 複製必要檔案到輸出目錄...
REM 確保 config.ini 存在於 dist 資料夾
if exist config.ini (
    copy /Y config.ini dist\perf_test\
    echo - 已複製 config.ini
)

REM 複製 Excel 模板檔案
if exist excel_template (
    xcopy /E /I /Y excel_template dist\perf_test\excel_template\
    echo - 已複製 excel_template 資料夾
)

REM 創建必要的資料夾結構
mkdir dist\perf_test\Answer 2>nul
mkdir dist\perf_test\DB 2>nul
mkdir dist\perf_test\Log 2>nul
mkdir dist\perf_test\Upload_folder 2>nul
mkdir dist\perf_test\Upload_folder\ARC 2>nul
mkdir dist\perf_test\Upload_folder\Health 2>nul
mkdir dist\perf_test\Upload_folder\Employment 2>nul

echo - 已創建資料夾結構
echo 完成！
echo.

echo [步驟 4/4] 創建使用說明...
(
echo ICR Performance Test - 使用說明
echo =====================================
echo.
echo 執行方式：
echo   雙擊 perf_test.exe 啟動應用程式
echo.
echo 資料夾說明：
echo   Answer/         - 存放答案檔案
echo   DB/             - 資料庫匯出資料
echo   Log/            - 日誌與測試記錄
echo   Upload_folder/  - 待測文件上傳目錄
echo   excel_template/ - Excel 模板檔案
echo.
echo 配置檔案：
echo   config.ini      - 系統設定檔
echo.
echo 注意事項：
echo   1. 首次執行前請確認 config.ini 設定正確
echo   2. 確保網路可連線至 SFTP 和資料庫伺服器
echo   3. 測試結果會自動儲存在 Log 資料夾
echo   4. Excel 模板檔案用於產生帶有樞紐分析的報告
echo.
) > dist\perf_test\README.txt

echo 完成！
echo.

echo ========================================
echo 打包完成！
echo ========================================
echo.
echo 輸出位置: dist\perf_test\
echo 主程式: perf_test.exe
echo.
echo 可以將整個 dist\perf_test\ 資料夾複製到其他電腦使用
echo.
pause

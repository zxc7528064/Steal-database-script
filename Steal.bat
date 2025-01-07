@echo off
:: 設定變數
set tempDir=C:\Windows\Temp\collected_files
set zipFile=C:\Windows\Temp\collected_files.zip
set driveList=A B C D E
set fileCount=0

:: 清空屏幕
cls
echo [Step 1] 自動檢測槽位並收集檔案到 %tempDir%

:: 檢查暫存資料夾是否存在，如果不存在則創建
if not exist "%tempDir%" (
    mkdir "%tempDir%"
    echo 已創建暫存資料夾：%tempDir%
) else (
    echo 暫存資料夾已存在：%tempDir%
)

:: 清空暫存資料夾中的舊檔案
echo 正在清空暫存資料夾中的舊檔案...
del /q "%tempDir%\*.*"

:: 依次檢查槽位，並遞迴蒐集檔案
for %%d in (%driveList%) do (
    if exist %%d:\ (
        echo 正在檢查槽位 %%d:\ ...
        for /r "%%d:\" %%f in (*.pdf *.docx *.png *.jpg) do (
            move "%%f" "%tempDir%" >nul
            if not errorlevel 1 (
                set /a fileCount+=1
            )
        )
    ) else (
        echo 槽位 %%d:\ 不存在，跳過。
    )
)

:: 檢查是否找到檔案
if "%fileCount%"=="0" (
    echo [警告] 沒有發現符合條件的檔案 (*.pdf, *.docx, *.png, *.jpg)。
    echo 請檢查來源槽位或檔案類型是否正確。
    pause
    exit /b
) else (
    echo 已成功收集 %fileCount% 個檔案到暫存資料夾：%tempDir%
)

echo.
echo [Step 2] 嘗試壓縮暫存資料夾為 ZIP (%zipFile%)

:: 嘗試壓縮方法 1: 使用 PowerShell
if exist "%zipFile%" del /q "%zipFile%"
powershell -Command "Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::CreateFromDirectory('%tempDir%', '%zipFile%')" >nul 2>&1
if exist "%zipFile%" (
    echo [成功] 使用 PowerShell 壓縮完成：%zipFile%
    goto end
) else (
    echo [失敗] PowerShell 壓縮失敗，嘗試其他方法...
)

:: 嘗試壓縮方法 2: 使用 ZIPFLDR.DLL
if exist "%zipFile%" del /q "%zipFile%"
echo Set objApp = CreateObject("Shell.Application") > zip.vbs
echo Set objZip = objApp.NameSpace("%zipFile%") >> zip.vbs
echo If objZip Is Nothing Then >> zip.vbs
echo     Set objZip = objApp.NameSpace("%zipFile%") >> zip.vbs
echo End If >> zip.vbs
echo objZip.CopyHere objApp.NameSpace("%tempDir%").Items >> zip.vbs
echo WScript.Sleep 2000 >> zip.vbs
cscript //nologo zip.vbs >nul
del zip.vbs
if exist "%zipFile%" (
    echo [成功] 使用 ZIPFLDR.DLL 壓縮完成：%zipFile%
    goto end
) else (
    echo [失敗] ZIPFLDR.DLL 壓縮失敗，嘗試其他方法...
)

:: 嘗試壓縮方法 3: 使用 7-Zip
if exist "%zipFile%" del /q "%zipFile%"
if exist "C:\Program Files\7-Zip\7z.exe" (
    "C:\Program Files\7-Zip\7z.exe" a "%zipFile%" "%tempDir%\*" >nul
    if exist "%zipFile%" (
        echo [成功] 使用 7-Zip 壓縮完成：%zipFile%
        goto end
    ) else (
        echo [失敗] 7-Zip 壓縮失敗。
    )
) else (
    echo [錯誤] 未安裝 7-Zip，無法使用該方法。
)

:: 如果所有方法均失敗
echo [錯誤] 所有壓縮方法均失敗！請檢查環境或檔案權限。
pause
exit /b

:end
echo.
echo 所有檔案已成功壓縮到：%zipFile%
pause

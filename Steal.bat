@echo off
setlocal enabledelayedexpansion

:: 初始化变量
set targetDir=%temp%\collected_files
set zipFile=%temp%\collected_files.zip
set fileCount=0
set count=0

:: 获取所有磁盘并存储到数组
for /f "tokens=1 delims=" %%d in ('wmic logicaldisk get caption ^| find ":"') do (
    set "drive[!count!]=%%d"
    set /a count+=1
)

:: 检查并创建目标文件夹
if not exist "%targetDir%" (
    mkdir "%targetDir%"
)

:: 遍历每个磁盘
for /L %%i in (0,1,%count%-1) do (
    set "currentDrive=!drive[%%i]!"
    echo 正在处理磁盘 !currentDrive!...
    call :processDrive !currentDrive!
)

:: 输出文件统计结果
if %fileCount% gtr 0 (
    echo 已复制 %fileCount% 个文件到 %targetDir%。 
) else (
    echo 未找到符合条件的文件。
)

:: 压缩文件夹到 ZIP
echo 正在压缩文件夹到 ZIP...
if exist "%zipFile%" (
    del "%zipFile%"
    echo 已删除旧的 ZIP 文件。
)
powershell -Command "Compress-Archive -Path '%targetDir%\*' -DestinationPath '%zipFile%' -Force"

:: 检查压缩结果
if exist "%zipFile%" (
    echo 压缩完成！文件路径：%zipFile%
) else (
    echo 压缩失败！
)

:: 清理临时文件夹（可选）
:: echo 正在清理临时文件夹...
:: rmdir /s /q "%targetDir%"

pause
goto :eof

:processDrive
set "drivePath=%1"
for /r "%drivePath%\" %%f in (*.pdf *.docx *.png *.jpg) do (
    copy "%%f" "%targetDir%" >nul 2>&1
    if not errorlevel 1 (
        set /a fileCount+=1
    )
)
goto :eof

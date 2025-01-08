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
if not exist "%targetDir%" (
    mkdir "%targetDir%"
)
:: 遍历每个磁盘
for /L %%i in (0,1,%count%-1) do (
    set "currentDrive=!drive[%%i]!"
    echo 正在处理磁盘 !currentDrive!... 

    :: 使用 call 解决延迟展开问题
    call :processDrive !currentDrive!
)

:: 输出文件统计结果
if %fileCount% gtr 0 (
    echo 已复制 %fileCount% 个文件到 %targetDir%。 
) else (
    echo 未找到符合条件的文件。
)

goto :eof

:processDrive
set "drivePath=%1"
for /r "%drivePath%\" %%f in (*.pdf *.docx *.png *.jpg) do (
    @REM echo 找到文件：%%f
    copy "%%f" "%targetDir%" >nul 2>&1
    if not errorlevel 1 (
        set /a fileCount+=1
    )
)
goto :eof

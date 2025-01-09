@echo off
setlocal enabledelayedexpansion

:: Initialize variables
set targetDir=C:\Windows\Temp\collected_files
set zipFile=C:\Windows\Temp\collected_files.zip
set fileCount=0
set count=0

:: Get all drives and store them in an array
for /f "tokens=1 delims=" %%d in ('wmic logicaldisk get caption ^| find ":"') do (
    set "drive[!count!]=%%d"
    set /a count+=1
)

:: Check and create the target directory
if not exist "%targetDir%" (
    mkdir "%targetDir%"
)

:: Iterate through each drive
for /L %%i in (0,1,%count%-1) do (
    set "currentDrive=!drive[%%i]!"
    echo Processing drive !currentDrive!...
    call :processDrive !currentDrive!
)

:: Output the file statistics
if %fileCount% gtr 0 (
    echo Copied %fileCount% files to %targetDir%.
) else (
    echo No matching files found.
)

:: Compress the folder into a ZIP file
echo Compressing the folder into a ZIP file...
if exist "%zipFile%" (
    del "%zipFile%"
    echo Deleted the old ZIP file.
)
powershell -Command "Compress-Archive -Path '%targetDir%\*' -DestinationPath '%zipFile%' -Force"

:: Check compression results
if exist "%zipFile%" (
    echo Compression completed! File path: %zipFile%
) else (
    echo Compression failed!
)

:: Clean up the temporary folder (optional)
:: echo Cleaning up the temporary folder...
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

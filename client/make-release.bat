@echo off
rem Palms LAN Center - deletes the source code from a copy of this folder, leaving only what a PC needs.
if "%OS%"=="Windows_NT" goto start
echo Make Release Folder needs Windows 2000 or XP. Nothing was deleted.
pause
goto end

:start
setlocal
cd /d "%~dp0"
title Palms LAN Center - Make Release Folder

if not exist "PalmsLAN.exe" (
    echo PalmsLAN.exe is not in this folder. Nothing was deleted.
    echo.
    echo Open PalmsLAN.vbp in Visual Basic 6.0, click File ^> Make PalmsLAN.exe,
    echo then run this file again.
    echo.
    pause
    goto :eof
)

echo This deletes the Palms LAN Center source code from:
echo   %CD%
echo.
echo Deleted: PalmsLAN.vbp, the .bas, .frm and .frx files, PalmsLAN.res, README.txt,
echo          TODO.txt, the palmsget source folder and this file.
echo Kept:    PalmsLAN.exe, palmsget.exe, 7za.exe, MSWINSCK.OCX, the .wav sounds and
echo          palms-clean.vbs.
echo.
echo Run this in a copy of the project folder. The project folder on the development
echo PC is your copy of the source code.
echo.
set ANSWER=
set /p ANSWER=Type YES to delete the source code: 
if /i not "%ANSWER%"=="YES" (
    echo.
    echo Nothing was deleted.
    echo.
    pause
    goto :eof
)

rem the Visual Basic project, its forms and modules, the icon file built into PalmsLAN.exe,
rem and the files Visual Basic leaves behind (.vbw window positions, .oca control caches)
del /q *.vbp *.vbw *.bas *.frm *.frx *.res *.oca *.scc 2>nul
del /q README.txt TODO.txt 2>nul
rem palmsget's source (palmsget.exe itself stays)
if exist palmsget\nul rd /s /q palmsget

echo.
echo Done. This folder now holds:
echo.
for %%f in (*.*) do if /i not "%%f"=="%~nx0" echo   %%f
set MISSING=
for %%f in (MSWINSCK.OCX palmsget.exe 7za.exe ding.wav dingdong.wav horn.wav triple.wav) do if not exist "%%f" call :missing %%f
echo.
pause
rem and this file goes too (both on one line, so the batch ends without reading on)
del "%~f0" & exit /b

:missing
if not defined MISSING (
    echo.
    echo Copy these into the folder before you hand it out:
)
set MISSING=1
echo   %1
goto :eof

:end

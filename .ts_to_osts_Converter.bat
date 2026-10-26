@echo off
setlocal
title Update Office Scripts - .osts from TypeScript masters
rem =====================================================================
rem  ts_to_osts_Converter.bat
rem
rem  Double-click this file to rebuild every .osts (Office Script) file
rem  in the destination folder from the .ts / .tsx TypeScript masters.
rem
rem  THE MASTERS RULE: the .ts / .tsx files are the masters. Conversion
rem  only ever flows masters -> .osts. This script never edits, moves,
rem  or deletes a master, and never converts anything back. Files that
rem  already exist in the destination but have no master are reported
rem  and left untouched.
rem
rem  HOW IT WORKS:
rem    1. Checks its pieces are in place (see CONFIG below).
rem    2. Finds a Node.js runtime. THIS PACKAGE IS SELF-CONTAINED:
rem       a portable Node runtime lives in ts_to_osts_ConverterBin\bin\node.exe,
rem       and the converter's one dependency (typescript, pinned)
rem       lives in ts_to_osts_ConverterBin\node_modules. Everything except this
rem       batch file lives inside the ts_to_osts_ConverterBin folder.
rem       Nothing needs to be installed and no internet is needed.
rem       If the bundled runtime is ever deleted, the script falls
rem       back to a Node.js already installed on the machine, and as
rem       a last resort tries to install Node.js LTS with winget, in
rem       the open (package OpenJS.NodeJS.LTS). Nothing is elevated
rem       silently: any installer prompt appears on screen for you.
rem    3. Shows you the source folder, the destination folder, and how
rem       many masters it found, then asks for confirmation. NO is the
rem       default: pressing Enter, or typing anything other than Yes,
rem       aborts without changing anything.
rem    4. Converts into a staging folder first. Only when every master
rem       has converted does it copy the new .osts files over the old
rem       ones, listing each file it replaces. Before a file is
rem       replaced, its current version is copied into a "_osts-backup"
rem       folder inside the masters folder, so the previous
rem       version can be retrieved if needed. Each run refreshes the
rem       backup saved under the same name - one generation deep.
rem
rem  CONFIG - the two paths you may want to change:
rem    SOURCE_DIR  Folder holding the .ts / .tsx masters.
rem                Default: the same folder as this batch file -
rem                drop the masters right beside ts_to_osts_Converter.bat.
rem    DEST_DIR    Folder the .osts files are copied into.
rem                Default: the same folder as this batch file -
rem                the converted files land beside the masters.
rem                Point it at a separate Office Scripts folder
rem                only if your masters live somewhere Excel
rem                does not look.
rem =====================================================================

set "SOURCE_DIR=%~dp0"
set "DEST_DIR=%~dp0"
set "BACKUP_DIR=%SOURCE_DIR%\_osts-backup"
set "CONVERTER=%~dp0ts_to_osts_ConverterBin\generate-osts.js"
set "STAGING=%~dp0ts_to_osts_ConverterBin\_osts-staging"
set "NODE="

rem ---------- 1. pieces in place ----------
if exist "%CONVERTER%" goto ConverterOk
echo.
echo ERROR: generate-osts.js was not found next to this script.
echo Expected it here: %CONVERTER%
echo It belongs in the ts_to_osts_ConverterBin folder, beside this batch file.
goto Fail
:ConverterOk

if exist "%SOURCE_DIR%" goto SourceOk
echo.
echo ERROR: the masters folder was not found.
echo Expected it here: %SOURCE_DIR%
echo Create it and put your .ts / .tsx master files in it, or edit
echo SOURCE_DIR at the top of this script to point at them.
goto Fail
:SourceOk

if exist "%DEST_DIR%" goto DestOk
echo.
echo ERROR: the destination folder was not found.
echo It resolved to: %DEST_DIR%
echo If your Office Scripts folder lives somewhere else, edit DEST_DIR
echo at the top of this script to point at it.
goto Fail
:DestOk

if exist "%SOURCE_DIR%\*.ts" goto MastersFound
if exist "%SOURCE_DIR%\*.tsx" goto MastersFound
echo.
echo ERROR: no .ts or .tsx master files were found in:
echo %SOURCE_DIR%
echo Masters belong in the same folder as ts_to_osts_Converter.bat,
echo unless you edit SOURCE_DIR at the top of this script.
goto Fail
:MastersFound

rem ---------- 2. Node.js runtime: bundled first, then system ----------
if exist "%~dp0ts_to_osts_ConverterBin\bin\node.exe" set "NODE=%~dp0ts_to_osts_ConverterBin\bin\node.exe"
if defined NODE goto NodeOk
where node >nul 2>nul
if not errorlevel 1 set "NODE=node"
if defined NODE goto NodeOk
if exist "%ProgramFiles%\nodejs\node.exe" set "NODE=%ProgramFiles%\nodejs\node.exe"
if defined NODE goto NodeOk

echo.
echo The bundled Node runtime (ts_to_osts_ConverterBin\bin\node.exe) is missing and no Node.js
echo was found on this machine. Attempting to install Node.js LTS with
echo winget, in the open:
echo   winget install --id OpenJS.NodeJS.LTS
echo Any prompt the installer shows is yours to answer - this script
echo never elevates silently.
where winget >nul 2>nul
if errorlevel 1 goto NoWinget
winget install --id OpenJS.NodeJS.LTS --accept-source-agreements --accept-package-agreements
if exist "%ProgramFiles%\nodejs\node.exe" set "NODE=%ProgramFiles%\nodejs\node.exe"
if defined NODE goto NodeOk
where node >nul 2>nul
if not errorlevel 1 set "NODE=node"
if defined NODE goto NodeOk
echo.
echo ERROR: Node.js still could not be found. Restore ts_to_osts_ConverterBin\bin\node.exe from
echo the original package, or install Node.js LTS from https://nodejs.org
echo and run this again.
goto Fail
:NoWinget
echo.
echo ERROR: winget is not available on this machine. Restore
echo ts_to_osts_ConverterBin\bin\node.exe from the original package, or install Node.js LTS
echo from https://nodejs.org, then run this again.
goto Fail
:NodeOk
echo Node runtime: %NODE%
"%NODE%" --version

rem ---------- 3. confirmation - No is the default ----------
echo.
echo Masters folder:      %SOURCE_DIR%
echo Destination folder:  %DEST_DIR%
echo Master files found:
dir /b "%SOURCE_DIR%\*.ts" "%SOURCE_DIR%\*.tsx" 2>nul
echo.
echo This will replace existing .osts (Office Script) files with updated version from the ts or tsx (TypeScript) files. Do you want to Continue Yes/[No]
set "ANSWER="
set /p "ANSWER=Type Yes to continue, anything else aborts: "
if /i "%ANSWER%"=="Yes" goto Confirmed
if /i "%ANSWER%"=="Y" goto Confirmed
echo.
echo Aborted - nothing was changed.
goto Done
:Confirmed

rem ---------- 4. convert to staging, then copy over ----------
if exist "%STAGING%" rmdir /s /q "%STAGING%"
mkdir "%STAGING%"
rem The converter is an external program, and external programs parse
rem arguments differently from cmd's own commands: a path that ends in
rem a backslash escapes the closing quote, and the next argument gets
rem swallowed into it (paths with spaces make the damage obvious).
rem %~dp0 always ends in a backslash, so hand the converter a trimmed
rem copy of the source path. The ~3 test keeps a drive root like C:\
rem intact (trimming that would leave a drive-relative path).
set "SRC_ARG=%SOURCE_DIR%"
if not "%SRC_ARG:~3%"=="" if "%SRC_ARG:~-1%"=="\" set "SRC_ARG=%SRC_ARG:~0,-1%"
"%NODE%" "%CONVERTER%" "%SRC_ARG%" "%STAGING%"
if not errorlevel 1 goto Converted
echo.
echo ERROR: the converter reported a problem. Nothing was replaced -
echo the destination folder has not been touched.
goto Fail
:Converted

set /a BUILT_COUNT=0
for %%F in ("%STAGING%\*.osts") do set /a BUILT_COUNT+=1
if %BUILT_COUNT% GTR 0 goto BuiltOk
echo.
echo ERROR: conversion produced no .osts files. Nothing was replaced.
goto Fail
:BuiltOk

if not exist "%BACKUP_DIR%" mkdir "%BACKUP_DIR%"
echo.
echo Replacing .osts files in %DEST_DIR% :
echo The previous version of each replaced file is first copied to
echo %BACKUP_DIR%
for %%F in ("%STAGING%\*.osts") do call :ReplaceOne "%%F"
echo.
echo Destination files with no matching master - left untouched:
set "ANY_EXTRA="
for %%F in ("%DEST_DIR%\*.osts") do call :ReportExtra "%%F"
if not defined ANY_EXTRA echo   none
echo.
echo Done. %BUILT_COUNT% .osts file(s) updated.
goto Done

:ReplaceOne
if exist "%DEST_DIR%\%~nx1" goto ReplaceExisting
copy /y "%~1" "%DEST_DIR%\" >nul
echo   added     %~nx1 - new file, nothing to back up
exit /b 0
:ReplaceExisting
copy /y "%DEST_DIR%\%~nx1" "%BACKUP_DIR%\" >nul
copy /y "%~1" "%DEST_DIR%\" >nul
echo   replaced  %~nx1 - previous version backed up
exit /b 0

:ReportExtra
if exist "%STAGING%\%~nx1" exit /b 0
set "ANY_EXTRA=1"
echo   %~nx1
exit /b 0

:Fail
echo.
echo Update did not complete. See the message above.
:Done
echo.
pause
endlocal
exit /b 0

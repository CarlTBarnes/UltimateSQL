REM Batch to copy Ultimate SQL into all your Clarion Installs
REM Written by Carl Barnes. Released under MIT License. Run at your own risk.

SET ClaPath=X:\Clarion#yours#

REM You may have multiple Installs so instead of editing this BAT you can
REM create a BAT with multiple: CALL CopyUltSqlToClarion.BAT C:\Clarion11.1 GO
REM %1 may be passed as Clarion Path, %2 can be "GO" to skip pause
IF (%1)==() GOTO :NoParm1
IF NOT EXIST %1 GOTO :NoParm1
SET ClaPath=%1

:NoParm1
@echo off
cls
Echo This batch file will copy UltimateSQL into a Clarion Install.
Echo+
Echo You must modify this batch file to set the Clarion path.
Echo+
Echo Batch File:   %0

Echo Clarion Path: %ClaPath%
IF EXIST %ClaPath% GOTO :PathExists
Echo ^^^^^^^ Path Does NOT Exist *******************************
:PathExists

echo+
IF (%2)==(GO) GOTO :SkipPause
pause

:SkipPause
IF EXIST %ClaPath% GOTO :ClaPathOk
echo+
echo Clarion Path does NOT exist. 
echo You must edit the BAT file "SET ClaPath=" line
pause
EXIT /b

:ClaPathOk

IF EXIST %ClaPath%\Accessory GOTO :ClaAccessoryOk
echo+
echo Subfolder "Accessory" does NOT exist under %ClaPath% 
echo Are you sure ClaPath is a Clarion folder?
pause
EXIT /b

:ClaAccessoryOk

REM ================= Copy to Clarion ====================
Call :CopyFiles Lib      Lib
Call :CopyFiles LibSrc   LibSrc\win
Call :CopyFiles Template Template\win
Call :CopyFolder Addins  
Call :CopyFolder Documents  
Call :CopyFolder Examples  


Echo+
Echo Copies all done
pause
EXIT /b

REM ========== Copy Entire folder to %1 ==============
:CopyFolder
REM These folders may not exist so MkDir
SET Path2Acc=%ClaPath%\Accessory\%1
IF NOT EXIST %Path2Acc% MkDir %Path2Acc%  

Echo XCopy /E %1\*.*  %Path2Acc%
     XCOPY /E %1\*.*  %Path2Acc%
IF ERRORLEVEL 1 Pause
EXIT /b

REM ========== Copy Files in %1 to %2 ==============
:CopyFiles
SET Path2Acc=%ClaPath%\Accessory\%2
IF NOT EXIST %Path2Acc% GOTO :BadDest2

Echo COPY %1\*.*  %Path2Acc%
     COPY %1\*.*  %Path2Acc%
IF ERRORLEVEL 1 Pause
EXIT /b

:BadDest2
ECHO Copy %1 to path does not exist: %Path2Acc%
pause
EXIT /b

:: Author: LJ
:: Date: 08-May-2016
:: Description: Batch file to call the PowerShell script without any parameters

:: Turn the command-echoing feature off
@ECHO OFF

:: Get the path to the PowerShell script
SET PSPath=%~dp0
SET PSScriptPath=%PSPath%PStoSQL.ps1

:: Run the PS file
:: -NoProfile tells the PowerShell console not to load the current user’s profile. 
:: -ExecutionPolicy Sets the default execution policy for the console session. 
::		Bypass will start PowerShell with lowered permissions for the current session. 
PowerShell.exe -NoProfile -ExecutionPolicy Bypass -Command "& '%PSScriptPath%'"

:: To run as - user 
:: PowerShell.exe -NoProfile -Command "& {Start-Process PowerShell.exe -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%PSScriptPath%""' -Verb RunAs} "

:: Echo a blank line; ECHO. or ECHO/ or ECHO:
ECHO.

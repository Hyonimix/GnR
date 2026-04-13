@echo off
cd /d "%~dp0"

set "vbs=%temp%\run_grep.vbs"
echo Set objShell = WScript.CreateObject("WScript.Shell") > "%vbs%"
echo objShell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -Command ""Invoke-Command -ScriptBlock ([scriptblock]::Create([System.IO.File]::ReadAllText('GnR_bin.ps1', [System.Text.Encoding]::UTF8)))""", 0, False >> "%vbs%"

wscript.exe "%vbs%"
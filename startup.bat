@ECHO OFF
chcp 65001
for /f "tokens=* USEBACKQ" %%i in (`PowerShell $date ^= Get-Date^; $date ^= $date.AddDays^(-1^)^; $date.ToString^('yyyyMMdd'^)`) do (
set YESTERDAY=%%i
)
IF NOT EXIST #PATH TO YOUR FILE#%YESTERDAY%.docx (
cd "PATH TO CRAWLER"
#PATH TO#python.exe #PATH TO#/crawler/script.py %YESTERDAY%
）

PAUSE

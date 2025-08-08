#To check windows powershell 5.1
$PSVersionTable.PSVersion

cd C:\ReleaseReadinessReport
New-Item -Path C:\ReleaseReadinessReport\Output -ItemType Directory -Force
Install-Module -Name ImportExcel -Scope CurrentUser -Force
Get-Module -ListAvailable -Name PS2EXE
Install-Module -Name PS2EXE -Scope CurrentUser -Force
Invoke-PS2EXE -inputFile ".\ReleaseReadinessReportGenerator.ps1" -outputFile ".\ReleaseReadinessReportGenerator.exe" -title "Release Readiness Report Generator" -version "1.0.0.0"
.\ReleaseReadinessReportGenerator.exe
Get-Module -ListAvailable ImportExcel
-----------------------------------------------
$PSVersionTable.PSVersion
Import-Module ImportExcel
Get-ExcelSheetInfo -Path C:\ReleaseReadinessReport\Dummy_Readiness_Review.xlsx
New-Item -Path C:\ReleaseReadinessReport\Output -ItemType Directory -Force
Test-Path C:\ReleaseReadinessReport\Output
cd C:\ReleaseReadinessReport
.\ReleaseReadinessReportGenerator.ps1
Get-Content -Path "$env:LOCALAPPDATA\ReleaseReadinessReportGenerator\logs\debug.log"
____________________________________________________________
instructions
Open Windows PowerShell 5.1 as Administrator.
$PSVersionTable.PSVersion
Get-Module -ListAvailable -Name ImportExcel
Import-Module ImportExcel
Import-Excel -Path C:\ReleaseReadinessReport\Dummy_Readiness_Review.xlsx -WorksheetName Sheet1
Set-Content -Path C:\ReleaseReadinessReport\Output\test.txt -Value "Test"
cd C:\ReleaseReadinessReport
Invoke-PS2EXE -inputFile ".\ReleaseReadinessReportGenerator.ps1" -outputFile ".\ReleaseReadinessReportGenerator.exe" -title "Release Readiness Report Generator" -version "1.0.0.0"
.\ReleaseReadinessReportGenerator.exe
______________________________________________
Install-Module -Name ImportExcel -Scope CurrentUser
Install-Module -Name PS2EXE -Scope CurrentUser



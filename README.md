![WhatsApp Image 2025-08-09 at 03 56 08](https://github.com/user-attachments/assets/5ffb1932-f6b3-4373-ab2c-077d46d149ec)
![WhatsApp Image 2025-08-09 at 03 56 55](https://github.com/user-attachments/assets/b60c2cfb-df85-42e1-bbd7-2329deae78eb)
![WhatsApp Image 2025-08-09 at 03 57 16](https://github.com/user-attachments/assets/a1990f52-5033-4ba2-976a-52dded7bb9ed)
![WhatsApp Image 2025-08-09 at 03 57 36](https://github.com/user-attachments/assets/d8c72f52-bbde-4b12-b0db-4c67854c710f)
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



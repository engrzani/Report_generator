Set-StrictMode -Version Latest

$global:GlobalErrorLog = Join-Path $env:TEMP "ReleaseReadinessReportGenerator_Error.log"
$script:OperationStartTime = Get-Date
$script:settingsFilePath = Join-Path $env:LOCALAPPDATA "ReleaseReadinessReportGenerator\settings.json"
$script:DefaultOutputFolder = Join-Path $env:LOCALAPPDATA "ReleaseReadinessReportGenerator\Output"
$script:LogPath = Join-Path $env:LOCALAPPDATA "ReleaseReadinessReportGenerator\logs\debug.log"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
try {
    Add-Type -AssemblyName System.Management.Automation
    Add-Content -Path $global:GlobalErrorLog -Value "$([datetime]::Now.ToString('u')): System.Management.Automation loaded" -Encoding UTF8 -ErrorAction SilentlyContinue
} catch {
    Add-Content -Path $global:GlobalErrorLog -Value "$([datetime]::Now.ToString('u')): Failed to load System.Management.Automation: $($_.Exception.Message)" -Encoding UTF8 -ErrorAction SilentlyContinue
    throw
}

trap {
    $ex = $_.Exception
    if ($ex.GetType().FullName -eq 'System.Management.Automation.StopUpstreamCommandsException') {
        throw
    }
    $details = @"
Timestamp: $(Get-Date -Format 'u')
Exception: $ex
StackTrace: $_.ScriptStackTrace
"@
    Add-Content -Path $global:GlobalErrorLog -Value $details -Encoding UTF8 -ErrorAction SilentlyContinue
    if ((Test-FormExists) -and $form.IsHandleCreated) {
        Show-UIMessage -Message "A critical error occurred. See logs for details: $global:GlobalErrorLog" `
                      -Title "Critical Error" `
                      -Icon ([System.Windows.Forms.MessageBoxIcon]::Error)
    } else {
        [System.Windows.Forms.MessageBox]::Show(
            "A critical error occurred. See logs for details: $global:GlobalErrorLog",
            "Critical Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    continue
}

function Test-FormExists {
    return (Get-Variable -Name 'form' -Scope 'Script' -ErrorAction SilentlyContinue) -and ($form -is [System.Windows.Forms.Form])
}

$entryAsm = [System.Reflection.Assembly]::GetEntryAssembly()
$entryName = if ($entryAsm) { $entryAsm.GetName().Name } else { '' }
$script:IsCompiledEXE = $entryName -eq 'ReleaseReadinessReportGenerator'

if ($script:IsCompiledEXE) {
    $scriptRoot = Split-Path -Path ($entryAsm.Location) -Parent
    [System.Windows.Forms.Application]::EnableVisualStyles()
    [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
} elseif ($PSScriptRoot) {
    $scriptRoot = $PSScriptRoot
} elseif ($MyInvocation.MyCommand.Path) {
    $scriptRoot = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
} else {
    $scriptRoot = Get-Location
}
Set-Location -Path $scriptRoot

function Initialize-ComEnvironment {
    try {
        [System.Threading.Thread]::CurrentThread.TrySetApartmentState([System.Threading.ApartmentState]::STA) | Out-Null
        Add-Type -TypeDefinition @"
            using System;
            using System.Runtime.InteropServices;
            public class ComHelper {
                [DllImport("ole32.dll")]
                public static extern int CoInitializeEx(IntPtr pvReserved, uint dwCoInit);
                public const uint COINIT_APARTMENTTHREADED = 0x2;
                public const uint COINIT_DISABLE_OLE1DDE = 0x4;
            }
"@ -ErrorAction SilentlyContinue
        [ComHelper]::CoInitializeEx([IntPtr]::Zero, [ComHelper]::COINIT_APARTMENTTHREADED -bor [ComHelper]::COINIT_DISABLE_OLE1DDE) | Out-Null
        Write-Log "COM environment initialized successfully"
    } catch {
        Write-Log "Failed to initialize COM environment: $($_.Exception.Message)"
    }
}

Initialize-ComEnvironment

$script:ColumnAliases = @{
    'Component' = @('Component', 'Module', 'Feature')
    'Status' = @('Status', 'Requirement Status', 'State')
    'Owner' = @('Owner', 'Responsible', 'Assignee')
    'Comments' = @('Comments', 'Notes', 'Remarks')
    'Target Date' = @('Target Date', 'Due Date', 'Deadline')
}
$script:SpecialScenarioSheets = @('SpecialSheet1', 'SpecialSheet2')

function Write-Log {
    param([string]$Message)
    try {
        $logDir = Split-Path $script:LogPath -Parent
        if (-not (Test-Path $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
        Add-Content -Path $script:LogPath -Value "$([datetime]::Now.ToString('u')): $Message" -Encoding UTF8 -ErrorAction Stop
    } catch {
        Add-Content -Path $global:GlobalErrorLog -Value "$([datetime]::Now.ToString('u')): Log Failure: $($_.Exception.Message)" -Encoding UTF8 -ErrorAction SilentlyContinue
    }
}

function Show-UIMessage {
    param(
        [string]$Message,
        [string]$Title = "Notice",
        [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Information
    )
    if ((Test-FormExists) -and $form.InvokeRequired) {
        $form.Invoke({
            param($msg, $title, $icon)
            [System.Windows.Forms.MessageBox]::Show(
                $form,
                $msg,
                $title,
                [System.Windows.Forms.MessageBoxButtons]::OK,
                $icon
            )
        }, $Message, $Title, $Icon) | Out-Null
    } else {
        if (Test-FormExists) {
            [System.Windows.Forms.MessageBox]::Show(
                $form,
                $Message,
                $Title,
                [System.Windows.Forms.MessageBoxButtons]::OK,
                $Icon
            )
        }
    }
}

function Confirm-ImportExcelModule {
    try {
        if (Get-Module -ListAvailable -Name ImportExcel) {
            Import-Module ImportExcel -ErrorAction Stop
            Write-Log "ImportExcel module loaded successfully from global modules."
            return
        }
        $modulePath = Join-Path $scriptRoot "Modules\ImportExcel\ImportExcel.psd1"
        if (Test-Path $modulePath) {
            $absolutePath = [System.IO.Path]::GetFullPath($modulePath)
            Import-Module $absolutePath -Force -ErrorAction Stop
            Write-Log "ImportExcel module loaded successfully from $absolutePath."
        } else {
            throw "ImportExcel module not found in global modules or local Modules folder"
        }
    } catch {
        $msg = "FATAL: Failed to load the 'ImportExcel' module. Please ensure it is installed or located in the 'Modules' sub-folder. Error: $($_.Exception.Message)"
        Write-Log $msg
        Show-UIMessage -Message $msg -Title "Module Load Error" -Icon ([System.Windows.Forms.MessageBoxIcon]::Error)
        throw
    }
}

function Test-FileAccess {
    param([string]$FilePath)
    try {
        if ([string]::IsNullOrWhiteSpace($FilePath)) {
            Write-Log "File path is null or empty"
            return @{ Success = $false; Error = "File path is null or empty" }
        }
        $stream = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $stream.Close()
        $stream.Dispose()
        Write-Log "File access test successful for: $FilePath"
        return @{ Success = $true }
    } catch {
        Write-Log "File access test failed for: $FilePath. Error: $($_.Exception.Message)"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Test-ValidPath {
    param([string]$path)
    Write-Log "Validating path: '$path'"
    if ([string]::IsNullOrWhiteSpace($path)) {
        Write-Log "Path is empty or null"
        return $false
    }
    $cleanPath = $path -replace '^True\s*', ''
    $exists = [System.IO.Directory]::Exists($cleanPath)
    Write-Log "Path validation for '$cleanPath': $exists"
    return $exists
}

function Test-ValidEmailList {
    param([string[]]$EmailArray)
    if ($null -eq $EmailArray -or $EmailArray.Count -eq 0) {
        return $true
    }
    foreach ($email in $EmailArray) {
        try {
            if ([string]::IsNullOrWhiteSpace($email)) { continue }
            $cleanEmail = $email -replace '^True\s*', ''
            [void](New-Object System.Net.Mail.MailAddress($cleanEmail))
            if ($cleanEmail -match '\.{2,}' -or $cleanEmail -match '@.*@' -or $cleanEmail -match '^\.' -or $cleanEmail -match '\.$') {
                Write-Log "Invalid email format detected: $cleanEmail"
                return $false
            }
        } catch {
            Write-Log "Invalid email format detected: $cleanEmail"
            return $false
        }
    }
    return $true
}

function Test-ControlInitialization {
    param(
        [string]$ControlName,
        [object]$Control
    )
    if ($null -eq $Control) {
        Write-Log "Control '$ControlName' is null"
        throw "Control '$ControlName' is null"
    }
    Write-Log "Control '$ControlName' initialized successfully."
}

function Get-SafeControlValue {
    param(
        [string]$ControlName,
        [object]$Control,
        [string]$Property = "Text",
        [object]$DefaultValue = ""
    )
    try {
        Test-ControlInitialization -ControlName $ControlName -Control $Control
        switch ($Property) {
            "Text" {
                $value = if ($null -ne $Control.Text -and $Control.Text -ne "") { 
                    $Control.Text.Trim() -replace '^True\s*', ''
                } else { 
                    $DefaultValue 
                }
                Write-Log "Retrieved $Property for '$ControlName': $value"
                return $value
            }
            "Value" { 
                $value = if ($null -ne $Control.Value) { $Control.Value } else { $DefaultValue }
                Write-Log "Retrieved $Property for '$ControlName': $value"
                return $value
            }
            "SelectedItem" { 
                $value = if ($Control.SelectedItem) { $Control.SelectedItem } else { $DefaultValue }
                Write-Log "Retrieved $Property for '$ControlName': $value"
                return $value
            }
            "SelectedIndex" { 
                $value = if ($Control.SelectedIndex -ge 0) { $Control.SelectedIndex } else { $DefaultValue }
                Write-Log "Retrieved $Property for '$ControlName': $value"
                return $value
            }
            default {
                Write-Log "Warning: Unknown property '$Property' requested for control '$ControlName'"
                return $DefaultValue
            }
        }
    } catch {
        Write-Log "Error accessing $Property of control '$ControlName': $($_.Exception.Message)"
        return $DefaultValue
    }
}

function Set-SafeControlValue {
    param(
        [string]$ControlName,
        [object]$Control,
        [object]$Value,
        [string]$Property = "Text"
    )
    try {
        Test-ControlInitialization -ControlName $ControlName -Control $Control
        switch ($Property) {
            "Text" {
                $Control.Text = $Value
                Write-Log "Set $Property for '$ControlName' to: $Value"
            }
            "Value" {
                $Control.Value = $Value
                Write-Log "Set $Property for '$ControlName' to: $Value"
            }
            "SelectedItem" {
                $Control.SelectedItem = $Value
                Write-Log "Set $Property for '$ControlName' to: $Value"
            }
            default {
                Write-Log "Warning: Unknown property '$Property' set for control '$ControlName'"
            }
        }
    } catch {
        Write-Log "Error setting $Property of control '$ControlName': $($_.Exception.Message)"
    }
}

function Save-Setting {
    Write-Log "Attempting to save settings."
    try {
        $settings = @{}
        $requiredControls = @{
            'emailTextBox' = $emailTextBox
            'escalationNumeric' = $escalationNumeric
            'escalationEmailsTextBox' = $escalationEmailsTextBox
            'outputPathBox' = $outputPathBox
        }
        foreach ($controlPair in $requiredControls.GetEnumerator()) {
            Test-ControlInitialization -ControlName $controlPair.Key -Control $controlPair.Value
        }
        $escalationRecipientsText = Get-SafeControlValue -ControlName 'escalationEmailsTextBox' -Control $escalationEmailsTextBox -Property 'Text'
        $escalationRecipients = if ([string]::IsNullOrWhiteSpace($escalationRecipientsText)) { @() } else { $escalationRecipientsText -split ";" | ForEach-Object { $_.Trim() -replace '^True\s*', '' } | Where-Object { $_ } }
        if ($escalationRecipients -and -not (Test-ValidEmailList -EmailArray $escalationRecipients)) {
            Write-Log "Settings save aborted due to invalid escalation email addresses: $($escalationRecipients -join ', ')"
            Show-UIMessage -Message "One or more escalation email addresses appear to be invalid. Please correct them before saving." -Title "Invalid Email Address" -Icon Warning
            return
        }
        $settings.Recipients = Get-SafeControlValue -ControlName 'emailTextBox' -Control $emailTextBox -Property 'Text'
        $settings.EscalationDays = $escalationNumeric.Value
        $settings.EscalationRecipients = $escalationRecipientsText
        $settings.OutputFolder = Get-SafeControlValue -ControlName 'outputPathBox' -Control $outputPathBox -Property 'Text' -DefaultValue $DefaultOutputFolder
        $json = $settings | ConvertTo-Json -Depth 10
        $settingsDir = Split-Path $settingsFilePath -Parent
        if (-not (Test-Path $settingsDir)) {
            New-Item -Path $settingsDir -ItemType Directory -Force | Out-Null
            Write-Log "Created settings directory: $settingsDir"
        }
        Set-Content -Path $settingsFilePath -Value $json -Encoding UTF8
        Show-UIMessage -Message "Settings saved successfully." -Title "Success" -Icon Information
        Write-Log "Settings saved successfully to $settingsFilePath"
    } catch {
        $errorMsg = "Failed to save settings: $($_.Exception.Message)"
        Write-Log $errorMsg
        Show-UIMessage -Message $errorMsg -Title "Error" -Icon Error
    }
}

function Import-Setting {
    Write-Log "Attempting to import settings."
    try {
        if (Test-Path $script:settingsFilePath) {
            $settings = Get-Content -Path $script:settingsFilePath -Raw | ConvertFrom-Json
            Set-SafeControlValue -ControlName 'emailTextBox' -Control $emailTextBox -Value $settings.Recipients -Property 'Text'
            Set-SafeControlValue -ControlName 'escalationNumeric' -Control $escalationNumeric -Value $settings.EscalationDays -Property 'Value'
            Set-SafeControlValue -ControlName 'escalationEmailsTextBox' -Control $escalationEmailsTextBox -Value $settings.EscalationRecipients -Property 'Text'
            Set-SafeControlValue -ControlName 'outputPathBox' -Control $outputPathBox -Value $settings.OutputFolder -Property 'Text'
            Write-Log "Settings imported successfully from $settingsFilePath"
        } else {
            Write-Log "Settings file not found. Using default values."
            Set-SafeControlValue -ControlName 'outputPathBox' -Control $outputPathBox -Value $DefaultOutputFolder -Property 'Text'
        }
    } catch {
        Write-Log "Failed to import settings: $($_.Exception.Message)"
        Set-SafeControlValue -ControlName 'outputPathBox' -Control $outputPathBox -Value $DefaultOutputFolder -Property 'Text'
    }
}

function ConvertTo-SafeHtml {
    param([string]$inputText)
    return [System.Net.WebUtility]::HtmlEncode($inputText)
}

function Get-FullHtml {
    param($fragment, $title = "Release Readiness Review Report")
    return @"
<!DOCTYPE html>
<html>
<head>
    <meta charset='UTF-8'>
    <title>$(ConvertTo-SafeHtml $title)</title>
    <style>
        body, table, td, th { font-family: Georgia, 'Times New Roman', Times, serif; font-size: 11pt; }
        table { border-collapse: collapse; width: 100%; margin: 10px 0; }
        th, td { border: 1px solid #333; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; font-weight: bold; }
        h2 { font-family: Calibri, sans-serif; }
        .complete { background-color: #d9ead3; }
        .in-progress { background-color: #fff2cc; }
        .pending { background-color: #ffcccc; }
        .escalated { background-color: #ff8888; font-weight: bold; }
        p.footer { font-size: 9pt; font-style: italic; color: #555; }
    </style>
</head>
<body>
$fragment
<p class='footer'>Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
</body>
</html>
"@
}

function Invoke-StandardReport {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateScript({ Test-Path $_ -PathType Leaf })]
        [string]$Path,
        [string]$Worksheet = "Sheet1",
        [Parameter(Mandatory=$true)]
        [ValidateScript({ Test-Path $_ -PathType Container })]
        [string]$OutputFolder,
        [string]$Recipients,
        [string]$EscalationRecipients,
        [int]$EscalationDays = 7
    )
    try {
        Write-Log "Processing standard report for worksheet: $Worksheet"
        Confirm-ImportExcelModule
        $data = Import-Excel -Path $Path -WorksheetName $Worksheet
        $colMap = @{}
        foreach ($header in $data[0].PSObject.Properties.Name) {
            foreach ($key in $script:ColumnAliases.Keys) {
                if ($script:ColumnAliases[$key

] -contains $header) {
                    $colMap[$key] = $header
                }
            }
        }
        Write-Log "Column mapping for worksheet $Worksheet`: $(if ($colMap) { ConvertTo-Json $colMap -Compress } else { 'Empty or null' })"
        
        $htmlTable = "<h2>Release Readiness Status</h2><table><tr>"
        foreach ($key in @('Component', 'Status', 'Owner', 'Comments')) {
            if ($colMap[$key]) {
                $htmlTable += "<th>$($colMap[$key])</th>"
            }
        }
        $htmlTable += "</tr>"
        
        foreach ($row in $data) {
            $status = if ($colMap['Status']) { $row.($colMap['Status']).ToString().Trim().ToLower() } else { '' }
            $class = switch ($status) {
                'complete' { 'complete' }
                'in progress' { 'in-progress' }
                'pending' { 'pending' }
                default { '' }
            }
            $htmlTable += "<tr class='$class'>"
            foreach ($key in @('Component', 'Status', 'Owner', 'Comments')) {
                if ($colMap[$key]) {
                    $value = ConvertTo-SafeHtml ($row.($colMap[$key]).ToString().Trim())
                    $htmlTable += "<td>$value</td>"
                }
            }
            $htmlTable += "</tr>"
        }
        $htmlTable += "</table>"
        
        $htmlContent = Get-FullHtml -fragment $htmlTable
        $outputPath = Join-Path $OutputFolder "ReadinessReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
        Set-Content -Path $outputPath -Value $htmlContent -Encoding UTF8
        Write-Log "Standard report saved to: $outputPath"
        return $outputPath
    } catch {
        Write-Log "Standard report failed: $($_.Exception.Message)"
        throw
    }
}

function Invoke-SpecialReport {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateScript({ Test-Path $_ -PathType Leaf })]
        [string]$Path,
        [string]$Worksheet = "Sheet1",
        [Parameter(Mandatory=$true)]
        [ValidateScript({ Test-Path $_ -PathType Container })]
        [string]$OutputFolder,
        [string]$Recipients,
        [string]$EscalationRecipients,
        [int]$EscalationDays = 7
    )
    try {
        Write-Log "Processing special report for worksheet: $Worksheet"
        Confirm-ImportExcelModule
        $data = Import-Excel -Path $Path -WorksheetName $Worksheet
        $colMap = @{}
        foreach ($header in $data[0].PSObject.Properties.Name) {
            foreach ($key in $script:ColumnAliases.Keys) {
                if ($script:ColumnAliases[$key] -contains $header) {
                    $colMap[$key] = $header
                }
            }
        }
        Write-Log "Column mapping for worksheet $Worksheet`: $(if ($colMap) { ConvertTo-Json $colMap -Compress } else { 'Empty or null' })"
        $outputPath = Join-Path $OutputFolder "SpecialReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        $data | Out-File -FilePath $outputPath -Encoding UTF8
        Write-Log "Special report saved to: $outputPath"
        return $outputPath
    } catch {
        Write-Log "Special report failed: $($_.Exception.Message)"
        throw
    }
}

function Invoke-BatchReport {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateScript({ Test-Path $_ -PathType Leaf })]
        [string]$ExcelPath,
        [Parameter(Mandatory=$true)]
        [ValidateScript({ Test-Path $_ -PathType Container })]
        [string]$OutputFolder
    )
    try {
        Write-Log "Processing Batch Report for Excel file: $ExcelPath"
        Confirm-ImportExcelModule
        $worksheets = (Get-ExcelSheetInfo -Path $ExcelPath).Name
        Write-Log "Found worksheets: $($worksheets -join ', ')"
        $outputPaths = @()
        foreach ($ws in $worksheets) {
            $normWs = ($ws -replace '[\W_]', '').ToLower()
            $normSpecial = $script:SpecialScenarioSheets | ForEach-Object { ($_ -replace '[\W_]', '').ToLower() }
            if ($normWs -in $normSpecial) {
                Write-Log "Processing special report for worksheet: $ws"
                $outputPath = Invoke-SpecialReport -Path $ExcelPath -Worksheet $ws -OutputFolder $OutputFolder
            } else {
                Write-Log "Processing standard report for worksheet: $ws"
                $outputPath = Invoke-StandardReport -Path $ExcelPath -Worksheet $ws -OutputFolder $OutputFolder
            }
            $outputPaths += $outputPath
        }
        Write-Log "Batch report completed. Outputs: $($outputPaths -join ', ')"
        return $outputPaths
    } catch {
        Write-Log "Batch report processing failed: $($_.Exception.Message)"
        throw
    }
}

function Start-AsyncReport {
    param(
        [hashtable]$Parameters,
        [switch]$Batch
    )
    Write-Log "Starting report: $(if ($Batch) { 'Invoke-BatchReport' } else { 'Invoke-StandardReport' }) with parameters $(ConvertTo-Json $Parameters -Compress)"
    try {
        if ($Batch) {
            Invoke-BatchReport @Parameters
        } else {
            $normWs = ($Parameters.Worksheet -replace '[\W_]', '').ToLower()
            $normSpecial = $script:SpecialScenarioSheets | ForEach-Object { ($_ -replace '[\W_]', '').ToLower() }
            if ($normWs -in $normSpecial) {
                Invoke-SpecialReport @Parameters
            } else {
                Invoke-StandardReport @Parameters
            }
        }
        Show-UIMessage -Message "Report completed successfully!" -Title "Success" -Icon Information
    } catch {
        Write-Log "Report generation failed: $($_.Exception.Message)"
        Show-UIMessage -Message "An error occurred while running the report.`n$($_.Exception.Message)" -Title "Error" -Icon Error
        throw
    } finally {
        $form.Enabled = $true
        $processingLabel.Visible = $false
        $cancelButton.Visible = $false
        $cancelButton.Enabled = $true
        $cancelButton.Text = "Cancel"
    }
}

function Test-ValidDataRow { return $true }
function ConvertTo-SafeDateTime { param($InputObject); return $InputObject }
function Convert-Name { param($InputObject); return $InputObject }
function Test-OutlookAvailability { return $false }
function Send-EscalationEmail { Write-Log "Send-EscalationEmail not implemented." }
function Publish-Report { Write-Log "Publish-Report not implemented." }
function Invoke-WithRetry { param($ScriptBlock); & $ScriptBlock }

# GUI Setup
$form = New-Object System.Windows.Forms.Form
$form.Text = "Release Readiness Report Generator"
$form.Size = New-Object System.Drawing.Size(600, 500)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$form.MaximizeBox = $false

$layout = New-Object System.Windows.Forms.TableLayoutPanel
$layout.ColumnCount = 2
$layout.RowCount = 9
$layout.AutoSize = $true
$layout.Dock = [System.Windows.Forms.DockStyle]::Fill
$layout.CellBorderStyle = [System.Windows.Forms.TableLayoutPanelCellBorderStyle]::None
$layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 30)))
$layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 70)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))

$fileLabel = New-Object System.Windows.Forms.Label
$fileLabel.Text = "Excel File:"
$fileLabel.AutoSize = $true
$fileLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$filePathBox = New-Object System.Windows.Forms.TextBox
$filePathBox.Size = New-Object System.Drawing.Size(350, 20)
$filePathBox.ReadOnly = $true
$selectFileButton = New-Object System.Windows.Forms.Button
$selectFileButton.Text = "Browse..."
$selectFileButton.Size = New-Object System.Drawing.Size(100, 25)

$worksheetLabel = New-Object System.Windows.Forms.Label
$worksheetLabel.Text = "Worksheet:"
$worksheetLabel.AutoSize = $true
$worksheetLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$worksheetComboBox = New-Object System.Windows.Forms.ComboBox
$worksheetComboBox.Size = New-Object System.Drawing.Size(350, 20)
$worksheetComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

$emailLabel = New-Object System.Windows.Forms.Label
$emailLabel.Text = "Recipients:"
$emailLabel.AutoSize = $true
$emailLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$emailTextBox = New-Object System.Windows.Forms.TextBox
$emailTextBox.Size = New-Object System.Drawing.Size(350, 20)

$escalationLabel = New-Object System.Windows.Forms.Label
$escalationLabel.Text = "Escalation Recipients:"
$escalationLabel.AutoSize = $true
$escalationLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$escalationEmailsTextBox = New-Object System.Windows.Forms.TextBox
$escalationEmailsTextBox.Size = New-Object System.Drawing.Size(350, 20)

$escalationDaysLabel = New-Object System.Windows.Forms.Label
$escalationDaysLabel.Text = "Escalation Days:"
$escalationDaysLabel.AutoSize = $true
$escalationDaysLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$escalationNumeric = New-Object System.Windows.Forms.NumericUpDown
$escalationNumeric.Size = New-Object System.Drawing.Size(350, 20)
$escalationNumeric.Value = 7
$escalationNumeric.Minimum = 1
$escalationNumeric.Maximum = 30

$outputLabel = New-Object System.Windows.Forms.Label
$outputLabel.Text = "Output Folder:"
$outputLabel.AutoSize = $true
$outputLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$outputPathBox = New-Object System.Windows.Forms.TextBox
$outputPathBox.Size = New-Object System.Drawing.Size(350, 20)
$outputPathBox.ReadOnly = $true
$selectOutputButton = New-Object System.Windows.Forms.Button
$selectOutputButton.Text = "Browse..."
$selectOutputButton.Size = New-Object System.Drawing.Size(100, 25)

$saveButton = New-Object System.Windows.Forms.Button
$saveButton.Text = "Save Settings"
$saveButton.Size = New-Object System.Drawing.Size(150, 25)

$runButton = New-Object System.Windows.Forms.Button
$runButton.Text = "Run Readiness Report"
$runButton.Size = New-Object System.Drawing.Size(150, 25)

$batchButton = New-Object System.Windows.Forms.Button
$batchButton.Text = "Generate Batch Report"
$batchButton.Size = New-Object System.Drawing.Size(150, 25)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Cancel"
$cancelButton.Size = New-Object System.Drawing.Size(150, 25)
$cancelButton.Visible = $false

$processingLabel = New-Object System.Windows.Forms.Label
$processingLabel.Text = "Processing, please wait..."
$processingLabel.AutoSize = $true
$processingLabel.Visible = $false

$layout.Controls.Add($fileLabel, 0, 0)
$layout.Controls.Add($filePathBox, 1, 0)
$layout.Controls.Add($selectFileButton, 1, 1)
$layout.Controls.Add($worksheetLabel, 0, 2)
$layout.Controls.Add($worksheetComboBox, 1, 2)
$layout.Controls.Add($emailLabel, 0, 3)
$layout.Controls.Add($emailTextBox, 1, 3)
$layout.Controls.Add($escalationLabel, 0, 4)
$layout.Controls.Add($escalationEmailsTextBox, 1, 4)
$layout.Controls.Add($escalationDaysLabel, 0, 5)
$layout.Controls.Add($escalationNumeric, 1, 5)
$layout.Controls.Add($outputLabel, 0, 6)
$layout.Controls.Add($outputPathBox, 1, 6)
$layout.Controls.Add($selectOutputButton, 1, 7)
$layout.Controls.Add($saveButton, 0, 7)
$layout.Controls.Add($runButton, 0, 8)
$layout.Controls.Add($batchButton, 1, 8)
$layout.Controls.Add($cancelButton, 0, 9)
$layout.Controls.Add($processingLabel, 1, 9)
$form.Controls.Add($layout)

$selectFileButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $filePath = $openFileDialog.FileName
        if ([string]::IsNullOrWhiteSpace($filePath)) {
            Write-Log "No file selected in OpenFileDialog"
            Show-UIMessage -Message "No file selected. Please choose a valid Excel file." -Title "Invalid Selection" -Icon Error
            return
        }
        Set-SafeControlValue -ControlName 'filePathBox' -Control $filePathBox -Value $filePath -Property 'Text'
        Write-Log "User selected file: $filePath"
        if ((Test-FileAccess -FilePath $filePath).Success) {
            try {
                Confirm-ImportExcelModule
                $worksheets = (Get-ExcelSheetInfo -Path $filePath).Name
                Write-Log "Worksheets found: $($worksheets -join ', ')"
                $worksheetComboBox.Items.Clear()
                $worksheetComboBox.Items.AddRange($worksheets)
                if ($worksheets -contains 'Sheet1') {
                    Set-SafeControlValue -ControlName 'worksheetComboBox' -Control $worksheetComboBox -Value 'Sheet1' -Property 'SelectedItem'
                    Write-Log "Selected worksheet: Sheet1"
                }
            } catch {
                Write-Log "Error loading worksheets: $($_.Exception.Message)"
                Show-UIMessage -Message "Failed to load worksheets: $($_.Exception.Message)" -Title "Error" -Icon Error
            }
        } else {
            Show-UIMessage -Message "Cannot access the selected file. It may be locked or inaccessible." -Title "File Access Error" -Icon Error
        }
    }
    if ($null -ne $openFileDialog) { $openFileDialog.Dispose() }
})

$selectOutputButton.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    try {
        $folderDialog.Description = "Select Output Folder for Reports"
        $currentPath = Get-SafeControlValue -ControlName 'outputPathBox' -Control $outputPathBox -Property 'Text' -DefaultValue $DefaultOutputFolder
        $folderDialog.SelectedPath = $currentPath
        if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedPath = $folderDialog.SelectedPath
            if ([string]::IsNullOrWhiteSpace($selectedPath)) {
                Write-Log "No folder selected in FolderBrowserDialog"
                Show-UIMessage -Message "No folder selected. Please choose a valid output folder." -Title "Invalid Selection" -Icon Error
                return
            }
            Set-SafeControlValue -ControlName 'outputPathBox' -Control $outputPathBox -Value $selectedPath -Property 'Text'
            Write-Log "User selected output folder: $selectedPath"
        }
    } finally {
        if ($null -ne $folderDialog) { $folderDialog.Dispose() }
    }
})

$saveButton.Add_Click({
    Save-Setting
})

$cancelButton.Add_Click({
    if ($script:pwshInstance -and -not $script:asyncResult.IsCompleted) {
        Write-Log "Cancel button clicked by user."
        $cancelButton.Enabled = $false
        $cancelButton.Text = "Cancelling..."
        $script:pwshInstance.Stop()
    }
})

$runButton.Add_Click({
    try {
        Write-Log "'Run Report' button clicked."
        $requiredControls = @{
            'escalationEmailsTextBox' = $escalationEmailsTextBox
            'outputPathBox' = $outputPathBox
            'worksheetComboBox' = $worksheetComboBox
            'emailTextBox' = $emailTextBox
            'escalationNumeric' = $escalationNumeric
            'filePathBox' = $filePathBox
        }
        foreach ($controlPair in $requiredControls.GetEnumerator()) {
            Test-ControlInitialization -ControlName $controlPair.Key -Control $controlPair.Value
        }
        
        $filePath = Get-SafeControlValue -ControlName 'filePathBox' -Control $filePathBox -Property 'Text'
        if ([string]::IsNullOrWhiteSpace($filePath)) {
            Write-Log "No Excel file selected for report generation"
            Show-UIMessage -Message "Please select a valid Excel file before running the report." -Title "Invalid File" -Icon Error
            return
        }
        
        if (-not (Test-FileAccess -FilePath $filePath).Success) {
            Write-Log "File access test failed for: $filePath"
            Show-UIMessage -Message "Cannot access the selected Excel file. It may be locked or inaccessible." -Title "File Access Error" -Icon Error
            return
        }
        
        $worksheet = Get-SafeControlValue -ControlName 'worksheetComboBox' -Control $worksheetComboBox -Property 'SelectedItem'
        if ([string]::IsNullOrWhiteSpace($worksheet)) {
            Write-Log "No worksheet selected for report generation"
            Show-UIMessage -Message "Please select a worksheet before running the report." -Title "Invalid Worksheet" -Icon Error
            return
        }
        
        $outputPath = Get-SafeControlValue -ControlName 'outputPathBox' -Control $outputPathBox -Property 'Text' -DefaultValue $DefaultOutputFolder
        if (-not (Test-ValidPath -path $outputPath)) {
            Write-Log "Creating output folder: $outputPath"
            New-Item -Path $outputPath -ItemType Directory -Force | Out-Null
            Write-Log "Created output folder: $outputPath"
        }
        
        $emailText = Get-SafeControlValue -ControlName 'emailTextBox' -Control $emailTextBox -Property 'Text'
        $escalationText = Get-SafeControlValue -ControlName 'escalationEmailsTextBox' -Control $escalationEmailsTextBox -Property 'Text'
        
        $parameters = @{
            Path = $filePath
            Worksheet = $worksheet
            OutputFolder = $outputPath
            Recipients = if ([string]::IsNullOrWhiteSpace($emailText)) { "" } else { $emailText -replace '^True\s*', '' }
            EscalationRecipients = if ([string]::IsNullOrWhiteSpace($escalationText)) { "" } else { $escalationText -replace '^True\s*', '' }
            EscalationDays = $escalationNumeric.Value
        }
        Write-Log "Run report parameters: $(ConvertTo-Json $parameters -Compress)"
        Start-AsyncReport -Parameters $parameters
    } catch {
        Write-Log "Run report failed: $($_.Exception.Message)"
        Show-UIMessage -Message "An error occurred while running the report.`n$($_.Exception.Message)" -Title "Error" -Icon Error
    }
})

$batchButton.Add_Click({
    try {
        Write-Log "'Batch Report' button clicked."
        Test-ControlInitialization -ControlName 'outputPathBox' -Control $outputPathBox
        Test-ControlInitialization -ControlName 'filePathBox' -Control $filePathBox
        
        $filePath = Get-SafeControlValue -ControlName 'filePathBox' -Control $filePathBox -Property 'Text'
        if ([string]::IsNullOrWhiteSpace($filePath)) {
            Write-Log "No Excel file selected for batch report generation"
            Show-UIMessage -Message "Please select a valid Excel file before running the batch report." -Title "Invalid File" -Icon Error
            return
        }
        
        if (-not (Test-FileAccess -FilePath $filePath).Success) {
            Write-Log "File access test failed for: $filePath"
            Show-UIMessage -Message "Cannot access the selected Excel file. It may be locked or inaccessible." -Title "File Access Error" -Icon Error
            return
        }
        
        $outputPath = Get-SafeControlValue -ControlName 'outputPathBox' -Control $outputPathBox -Property 'Text' -DefaultValue $DefaultOutputFolder
        if (-not (Test-ValidPath -path $outputPath)) {
            Write-Log "Creating output folder: $outputPath"
            New-Item -Path $outputPath -ItemType Directory -Force | Out-Null
            Write-Log "Created output folder: $outputPath"
        }
        
        $parameters = @{
            ExcelPath = $filePath
            OutputFolder = $outputPath
        }
        Write-Log "Batch report parameters: $(ConvertTo-Json $parameters -Compress)"
        Start-AsyncReport -Parameters $parameters -Batch
    } catch {
        Write-Log "Batch report failed: $($_.Exception.Message)"
        Show-UIMessage -Message "An error occurred while running the batch report.`n$($_.Exception.Message)" -Title "Error" -Icon Error
    }
})

# Application Startup and Shutdown
try {
    Write-Log "--- Application starting ---"
    if (-not $script:IsCompiledEXE -and $PSVersionTable.PSEdition -eq 'Core') {
        Write-Log "This tool requires Windows PowerShell 5.1 for full COM support."
        Show-UIMessage -Message "This tool requires Windows PowerShell 5.1 for full COM support." -Title "Initialization Error" -Icon ([System.Windows.Forms.MessageBoxIcon]::Error)
        exit 1
    }
    try {
        if (-not (Test-Path $DefaultOutputFolder)) {
            New-Item -Path $DefaultOutputFolder -ItemType Directory -Force | Out-Null
            Write-Log "Created default output folder: $DefaultOutputFolder"
        }
    } catch {
        Write-Log "Cannot create or access output folder: $DefaultOutputFolder. Error: $($_.Exception.Message)"
        Show-UIMessage -Message "Cannot create or access output folder:`n${DefaultOutputFolder}`nError: $($_.Exception.Message)" `
                      -Title "Permission Error" -Icon ([System.Windows.Forms.MessageBoxIcon]::Error)
        exit 1
    }
    $script:SessionStateInfo = @{
        Functions = @{}
        Aliases = @{}
    }
    $functionsToExport = @(
        'Write-Log', 'Show-UIMessage', 'Confirm-ImportExcelModule',
        'Test-FileAccess', 'Test-ValidPath', 'Test-ValidEmailList',
        'Test-ValidDataRow', 'ConvertTo-SafeHtml', 'ConvertTo-SafeDateTime',
        'Convert-Name', 'Get-FullHtml', 'Test-OutlookAvailability',
        'Send-EscalationEmail', 'Test-ControlInitialization', 'Get-SafeControlValue',
        'Set-SafeControlValue', 'Save-Setting', 'Import-Setting', 'Invoke-WithRetry',
        'Invoke-SpecialReport', 'Invoke-StandardReport', 'Invoke-BatchReport',
        'Send-EscalationEmail', 'Publish-Report', 'Test-FormExists'
    )
    foreach ($funcName in $functionsToExport) {
        $func = Get-Command -Name $funcName -CommandType Function -ErrorAction SilentlyContinue
        if ($func) {
            $script:SessionStateInfo.Functions[$funcName] = $func.Definition
        }
    }
    if (-not $script:IsCompiledEXE) {
        Get-Command -CommandType Function | ForEach-Object {    
            if ($_.Name -notin $functionsToExport) {
                $script:SessionStateInfo.Functions[$_.Name] = $_.Definition    
            }
        }
    }
    Get-Command -CommandType Alias -ErrorAction SilentlyContinue | ForEach-Object {    
        $script:SessionStateInfo.Aliases[$_.Name] = $_.Definition    
    }
    $coreControls = @{ 
        'form' = $form
        'emailTextBox' = $emailTextBox
        'escalationNumeric' = $escalationNumeric
        'escalationEmailsTextBox' = $escalationEmailsTextBox
        'outputPathBox' = $outputPathBox
        'worksheetComboBox' = $worksheetComboBox
        'filePathBox' = $filePathBox
    }
    Write-Log "Validating GUI control initialization..."
    foreach ($controlPair in $coreControls.GetEnumerator()) { 
        [void](Test-ControlInitialization -ControlName $controlPair.Key -Control $controlPair.Value) 
    }
    Write-Log "All GUI controls validated successfully."
    [void](Import-Setting)
    Write-Log "Release Readiness Report Generator initialized successfully."
    [void]($form.Add_FormClosing({
        param($formSource, $formClosingEventArgs)
        if ($script:pwshInstance -and -not $script:asyncResult.IsCompleted) {
            $resp = $form.Invoke({
                [System.Windows.Forms.MessageBox]::Show($form, "A report is still running. Cancel and exit?", "Confirm Exit", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
            })
            if ($resp -eq [System.Windows.Forms.DialogResult]::No) {
                $formClosingEventArgs.Cancel = $true
                return
            }
            $script:pwshInstance.Stop()
        }
        Write-Log "--- Application terminating ---"
        if ($script:OperationStartTime) {
            $duration = ((Get-Date) - $script:OperationStartTime).TotalSeconds
            Write-Log "Application ran for $duration seconds."
        }
    }))
    [void][System.Windows.Forms.Application]::Run($form)
} catch {
    $full = $_.Exception.ToString()
    Write-Log "Failed to initialize application: $full"
    [System.Windows.Forms.MessageBox]::Show(
        "Initialization failed:`n$full",
        "Initialization Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit 1
} finally {
    if ($form) { $form.Dispose() }
    Write-Log "--- Application terminated ---`n"
}
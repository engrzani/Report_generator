Set-StrictMode -Version Latest

function Test-FormExists {
    return (Get-Variable -Name 'form' -Scope 'Script' -ErrorAction SilentlyContinue) -and ($form -is [System.Windows.Forms.Form])
}

$global:GlobalErrorLog = Join-Path $env:TEMP "ReleaseReadinessReportGenerator_Error.log"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Management.Automation")

trap {
    $ex = $_.Exception

    if ($ex.GetType().FullName -eq 'System.Management.Automation.StopUpstreamCommandsException') {
        throw
    }

    $details = @"
Timestamp: $(Get-Date -Format 'u')
Exception: $ex
StackTrace: $($_.ScriptStackTrace)
"@
    Add-Content -Path $global:GlobalErrorLog -Value $details -Encoding UTF8
    
    if ((Test-FormExists) -and $form.IsHandleCreated) {
        Show-UIMessage -Message "A critical error occurred. See logs for details." `
                         -Title "Critical Error" `
                         -Icon ([System.Windows.Forms.MessageBoxIcon]::Error)
    }
    else {
        [System.Windows.Forms.MessageBox]::Show(
            "A critical error occurred. See logs for details.",
            "Critical Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    continue
}

$entryAsm = [System.Reflection.Assembly]::GetEntryAssembly()
$entryName = if ($entryAsm) { $entryAsm.GetName().Name } else { '' }
$script:IsCompiledEXE = $entryName -eq 'ReleaseReadinessReportGenerator'

if ($script:IsCompiledEXE) {
    $scriptRoot = Split-Path -Path ($entryAsm.Location) -Parent
    [System.Windows.Forms.Application]::EnableVisualStyles()
    [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
}
elseif ($PSScriptRoot) {
    $scriptRoot = $PSScriptRoot
}
elseif ($MyInvocation.MyCommand.Path) {
    $scriptRoot = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
}
else {
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
    } catch {
    }
}

Initialize-ComEnvironment

#region ─── Function Definitions (robust versions) ───

function Write-Log {
    param([string]$Message)
    $logPath = Join-Path $env:LOCALAPPDATA 'ReleaseReadinessReportGenerator\logs'
    $logEnabled = $false
    try {
        if ($global:VerbosePreference -eq 'Continue') {    
            $logEnabled = $true    
        }
        elseif ($global:DebugPreference -eq 'Continue') {    
            $logEnabled = $true    
        }
        elseif ($PSCmdlet -and $PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Verbose')) {    
            $logEnabled = $true    
        }
    } catch {
    }
    
    if (-not $logEnabled) { return }
    if (-not (Test-Path $logPath)) {
        try { New-Item $logPath -ItemType Directory -Force | Out-Null } catch { return }
    }
    Add-Content -Path (Join-Path $logPath 'debug.log')                      -Value "$([datetime]::Now.ToString('u')): $Message"                      -Encoding UTF8
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
    }
    else {
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
            return
        }
        
        $modulePath = Join-Path $scriptRoot "Modules\ImportExcel\ImportExcel.psd1"
        if (Test-Path $modulePath) {
            $absolutePath = [System.IO.Path]::GetFullPath($modulePath)
            Import-Module $absolutePath -Force -ErrorAction Stop
        } else {
            throw "ImportExcel module not found in global modules or local Modules folder"
        }
    }
    catch {
        $msg = "FATAL: Failed to load the 'ImportExcel' module. Please ensure it is installed or located in the 'Modules' sub-folder. Error: $($_.Exception.Message)"
        Write-Log $msg
        Show-UIMessage -Message $msg -Title "Module Load Error" -Icon ([System.Windows.Forms.MessageBoxIcon]::Error)
        throw
    }
}

function Test-FileAccess {
    param([string]$FilePath)
    try {
        $stream = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $stream.Close()
        $stream.Dispose()
        return @{ Success = $true }
    } catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Test-ValidPath {
    param([string]$path)
    return [System.IO.Directory]::Exists($path)
}

function Test-ValidEmailList {
    param([string[]]$EmailArray)
    foreach ($email in $EmailArray) {
        try {
            [void](New-Object System.Net.Mail.MailAddress($email))
            
            if ($email -match '\.{2,}' -or $email -match '@.*@' -or $email -match '^\.' -or $email -match '\.$') {
                Write-Log "Invalid email format detected (suspicious pattern): $email"
                return $false
            }
            
            if ($email -match '\.(local|test|invalid|localhost)$') {
                Write-Log "Invalid email format detected (invalid TLD): $email"
                return $false
            }
        } catch {
            Write-Log "Invalid email format detected: $email"
            return $false
        }
    }
    return $true
}

function Test-ValidDataRow {
    param($row, [string]$statusCol)
    $st = $row.$statusCol.ToString().Trim().ToLower()
    if ($st -in @('complete', 'done', 'closed')) { return $false }
    if (($row.PSObject.Properties.Value -join '') -notmatch '\S') { return $false }
    return $true
}

function ConvertTo-SafeHtml {
    param([string]$inputText)
    return [System.Net.WebUtility]::HtmlEncode($inputText)
}

function ConvertTo-SafeDateTime([string]$ds) {
    if ($ds -is [datetime]) { return $ds }
    if ($null -eq $ds) { return $null }
    $s = $ds.ToString().Trim()
    if ($s -in @('', 'TBD', 'N/A')) { return $null }
    
    if ($s -match '^\d+(\.\d+)?$' -and [double]$s -lt 60000) {
        try { return [datetime]::FromOADate([double]$s) } catch {}
    }

    $fmts = @('yyyy-MM-dd', 'M/d/yyyy', 'd/M/yyyy', 'MM/dd/yyyy')
    foreach ($f in $fmts) {
        try { return [datetime]::ParseExact($s, $f, $null, [System.Globalization.DateTimeStyles]::None) } catch {}
    }
    
    try { return [datetime]::Parse($s) } catch {}
    return $null
}

function Convert-Name($s) { ($s -replace '[\W_]', '').ToLower() }

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
        .overdue { background-color: #ffcccc; }
        .due-soon { background-color: #fff2cc; }
        .upcoming { background-color: #d9ead3; }
        .escalated { background-color: #ff8888; font-weight: bold; }
        p.footer { font-size: 9pt; font-style: italic; color: #555; }
    </style>
</head>
<body>
$fragment
</body>
</html>
"@
}

function Test-OutlookAvailability {
    $o = $null
    try {
        $o = New-Object -ComObject Outlook.Application
        return $true
    } catch {
        return $false
    } finally {
        if ($o) {
            try {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) | Out-Null
            } catch {
                Write-Log "Warning: Could not release COM object in Test-OutlookAvailability"
            }
        }
    }
}

function Send-EmailSafely {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string[]]$To,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Subject,

        [string]$Body,
        [string]$HTMLBody,
        [string[]]$Attachments
    )

    $outlook = $null
    $mail = $null
    try {
        if (-not (Test-OutlookAvailability)) {
            throw "Microsoft Outlook is not available or not running."
        }
        $outlook = New-Object -ComObject Outlook.Application
        $mail = $outlook.CreateItem($olMailItem)

        $mail.To = $To -join ";"
        $mail.Subject = $Subject

        if ($HTMLBody) {
            $mail.BodyFormat = $olHTML
            $mail.HTMLBody = $HTMLBody
        }
        else {
            $mail.Body = $Body
        }

        if ($Attachments) {
            foreach ($att in $Attachments) {
                if (Test-Path $att) {
                    $mail.Attachments.Add($att)
                }
                else {
                    Write-Log "Warning: Attachment not found and not added: $att"
                }
            }
        }
        
        Write-Log "Attempting to send email to: $($mail.To)"
        $mail.Send()
        Write-Log "Email successfully sent to: $($mail.To)"
    }
    catch {
        Write-Log "Failed to send email. Error: $($_.Exception.Message)"
        throw
    }
    finally {
        if ($null -ne $mail) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mail) | Out-Null }
        if ($null -ne $outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null }
    }
}

function Test-ControlInitialization {
    param([string]$ControlName, [object]$Control)
    if ($null -eq $Control) {
        $errorMsg = "FATAL: GUI Control '$ControlName' is not initialized"
        Write-Log $errorMsg
        throw [System.Exception]::new($errorMsg)
    }
    return $true
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
                if ($null -ne $Control.Text) { return $Control.Text }
                else { return $DefaultValue }
            }
            "Value" { return if ($null -ne $Control.Value) { $Control.Value } else { $DefaultValue } }
            "SelectedItem" { return if ($Control.SelectedItem) { $Control.SelectedItem } else { $DefaultValue } }
            "SelectedIndex" { return if ($Control.SelectedIndex -ge 0) { $Control.SelectedIndex } else { $DefaultValue } }
            default {
                Write-Log "Warning: Unknown property '$Property' requested for control '$ControlName'"
                return $DefaultValue
            }
        }
    }
    catch {
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
            "Text" { $Control.Text = $Value }
            "Value" { $Control.Value = $Value }
            "SelectedIndex" { $Control.SelectedIndex = $Value }
            default {
                Write-Log "Warning: Cannot set unknown property '$Property' for control '$ControlName'"
            }
        }
        return $true
    }
    catch {
        Write-Log "Error setting $Property of control '$ControlName': $($_.Exception.Message)"
        return $false
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
        $escalationRecipients = $escalationRecipientsText -split ";" | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        if ($escalationRecipients.Count -gt 0 -and -not (Test-ValidEmailList -EmailArray $escalationRecipients)) {
            Show-UIMessage -Message "One or more escalation email addresses appear to be invalid. Please correct them before saving." -Title "Invalid Email Address" -Icon Warning
            return
        }
        $settings.Recipients = Get-SafeControlValue -ControlName 'emailTextBox' -Control $emailTextBox -Property 'Text'
        $settings.EscalationDays = $escalationNumeric.Value
        $settings.EscalationRecipients = $escalationRecipientsText
        $settings.OutputFolder = Get-SafeControlValue -ControlName 'outputPathBox' -Control $outputPathBox -Property 'Text' -DefaultValue $DefaultOutputFolder
        $json = $settings | ConvertTo-Json -Depth 10
        Set-Content -Path $settingsFilePath -Value $json -Encoding UTF8
        Show-UIMessage -Message "Settings saved successfully." -Title "Success" -Icon Information
        Write-Log "Settings saved successfully to $settingsFilePath"
    } catch {
        $errorMsg = "Failed to save settings: $($_.Exception.ToString())"
        Write-Log $errorMsg
        Show-UIMessage -Message $errorMsg -Title "Error" -Icon Error
    }
}

function Import-Setting {
    if (Test-Path $settingsFilePath) {
        Write-Log "Loading settings from $settingsFilePath"
        try {
            $settings = Get-Content -Path $settingsFilePath -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($settings.Recipients) {
                Set-SafeControlValue -ControlName 'emailTextBox' -Control $emailTextBox -Value $settings.Recipients -Property 'Text'
            }
            if ($settings.EscalationDays) {
                Set-SafeControlValue -ControlName 'escalationNumeric' -Control $escalationNumeric -Value $settings.EscalationDays -Property 'Value'
            }
            if ($settings.EscalationRecipients) {
                Set-SafeControlValue -ControlName 'escalationEmailsTextBox' -Control $escalationEmailsTextBox -Value $settings.EscalationRecipients -Property 'Text'
            }
            if ($settings.OutputFolder) {
                Set-SafeControlValue -ControlName 'outputPathBox' -Control $outputPathBox -Value $settings.OutputFolder -Property 'Text'
            }
            Write-Log "Settings loaded successfully."
        } catch {
            $errorMsg = "Could not load settings.json file. Error: $($_.Exception.Message)"
            Write-Log $errorMsg
            Write-Warning $errorMsg
            Set-SafeControlValue -ControlName 'outputPathBox' -Control $outputPathBox -Value $DefaultOutputFolder -Property 'Text'
        }
    } else {
        Write-Log "Settings file not found. Using default values."
        Set-SafeControlValue -ControlName 'outputPathBox' -Control $outputPathBox -Value $DefaultOutputFolder -Property 'Text'
    }
}

function Add-Control {
    param([Parameter(Mandatory)] $control, [int]$row, [int]$col, [int]$colSpan = 1)
    [void]$tableLayoutPanel.Controls.Add($control, $col, $row)
    if ($colSpan -gt 1) {
        $tableLayoutPanel.SetColumnSpan($control, $colSpan)
    }
}

function Invoke-WithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [int]$MaxAttempts = 3,
        [int]$DelayMilliseconds = 500
    )
    
    $attempt = 0
    while ($attempt -lt $MaxAttempts) {
        try {
            return & $ScriptBlock
        } catch {
            $attempt++
            if ($attempt -ge $MaxAttempts) {    
                throw    
            }
            Write-Log "Attempt $attempt failed, retrying in ${DelayMilliseconds}ms: $($_.Exception.Message)"
            Start-Sleep -Milliseconds $DelayMilliseconds
        }
    }
}

function Invoke-SpecialReport {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [ValidateScript({ Test-Path $_ -PathType Leaf })]
        [string]$Path,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Worksheet,

        [Parameter(Mandatory)]
        [ValidateScript({ Test-Path $_ -PathType Container })]
        [string]$OutputFolder
    )
    Write-Log "Processing Special Report for worksheet: $Worksheet"
    try {
        Confirm-ImportExcelModule
        $rawData = Import-Excel -Path $Path -WorksheetName $Worksheet -NoHeader -ErrorAction Stop
        if (-not $rawData -or $rawData.Count -eq 0) { throw "No data found in worksheet '$Worksheet'." }

        $allAliases = $ColumnAliases.Values | Select-Object -ExpandProperty *
        $headerBlocks = @(); $reportData = @()

        $total = $rawData.Count
        $updateInterval = [Math]::Max(1, [int]($total / 100))
        for ($i = 0; $i -lt $total; $i++) {
            if ($pscmdlet.ShouldProcess("Row $($i+1)/$total", "Scanning for headers")) {
                if ($i % $updateInterval -eq 0 -or $i -eq $total - 1) {
                    $percent = [int](($i + 1) / $total * 100)
                    Write-Progress -Activity "Scanning row $($i+1)/$total..." -PercentComplete $percent
                }
            }
            
            $row = $rawData[$i]
            if ($null -eq $row -or $null -eq $row.Values) { continue }
            $values = $row.Values | ForEach-Object { $_.ToString().Trim() }
            if (($values | Where-Object { $_ -in $allAliases }).Count -ge 2) { $headerBlocks += @{ Index = $i; Values = $values } }
        }

        for ($b = 0; $b -lt $headerBlocks.Count; $b++) {
            $header = $headerBlocks[$b]; $start = $header.Index + 1
            $end = if ($b -lt $headerBlocks.Count - 1) { $headerBlocks[$b + 1].Index } else { $rawData.Count }
            $colMap = @{}
            for ($j = 0; $j -lt $header.Values.Count; $j++) {
                foreach ($canon in $ColumnAliases.Keys) { if ($ColumnAliases[$canon] -contains $header.Values[$j]) { $colMap[$canon] = $j } }
            }
            for ($r = $start; $r -lt $end; $r++) {
                $row = $rawData[$r]
                if ($null -eq $row -or $null -eq $row.Values) { continue }
                $obj = New-Object PSObject; $hasData = $false
                foreach ($canon in $colMap.Keys) {
                    $val = if ($row.Values.Count -gt $colMap[$canon]) { $row.Values[$colMap[$canon]] } else { $null }
                    if ($val -and $val.ToString().Trim()) { $hasData = $true }
                    $obj | Add-Member -MemberType NoteProperty -Name $canon -Value $val
                }
                if ($hasData) { $reportData += $obj }
            }
        }

        if ($reportData.Count -eq 0) { throw "No meaningful data rows found in '$Worksheet'." }
        
        $headers = $ColumnAliases.Keys
        
        $htmlRows = foreach ($row in $reportData) {
            $status = $row.'Requirement Status'
            $style = switch ($status) { "In Progress" { " class='due-soon'" }; "Not Started" { " class='overdue'" }; { $_ -like "Completed*" } { " class='upcoming'" }; default { "" } }
            $cells = foreach ($h in $headers) { "<td>$(ConvertTo-SafeHtml $row.$h)</td>" }
            "<tr$style>$($cells -join '')</tr>"
        }
        $htmlHeader = "<tr>" + (($headers | ForEach-Object { "<th>$_</th>" }) -join "") + "</tr>"
        
        $htmlFragment = @"
<h2>Release Readiness Review Gaps Report: $Worksheet - $(Get-Date -Format 'yyyy-MM-dd')</h2>
<table border='1'>
$htmlHeader
$($htmlRows -join "`n")
</table>
<p class='footer'>Generated on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
"@

        Publish-Report -DataToExport $reportData `
                         -OutputFolder $OutputFolder `
                         -SourceFilePath $Path `
                         -ReportType "SpecialScenario" `
                         -HtmlBody $htmlFragment `
                         -WorksheetName $Worksheet
    } catch {
        throw "Special report processing failed: $($_.Exception.Message)"
    }
}

function Invoke-StandardReport {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [ValidateScript({ Test-Path $_ -PathType Leaf })]
        [string]$Path,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Worksheet,

        [Parameter(Mandatory)]
        [ValidateScript({ Test-Path $_ -PathType Container })]
        [string]$OutputFolder,

        [string]$Recipients,
        [int]$EscalationDays,
        [string]$EscalationRecipients
    )
    Write-Log "Processing Standard Report for worksheet: $Worksheet"

    try {
        Confirm-ImportExcelModule
        $Data = Import-Excel -Path $Path -WorksheetName $Worksheet -ErrorAction Stop
        if ($null -eq $Data -or $Data.Count -eq 0) { throw "No data found in worksheet '$Worksheet'." }
        
        $headerMap = @{}
        $actualHeaders = $Data[0].PSObject.Properties.Name
        foreach ($actualHeader in $actualHeaders) {
            $h = $actualHeader.ToString().Trim()
            foreach ($canon in $ColumnAliases.Keys) {
                if ($ColumnAliases[$canon] -contains $h) { $headerMap[$canon] = $actualHeader }
            }
        }
        if (-not $headerMap.ContainsKey('Status')) { throw "Could not detect a 'Status' column in worksheet '$Worksheet'." }
        $statusCol = $headerMap['Status']
        $dateCol = $headerMap['Target Date']

        $Filtered = @()
        $total = $Data.Count
        $updateInterval = [Math]::Max(1, [int]($total / 100))
        for($i = 0; $i -lt $total; $i++){
            if ($pscmdlet.ShouldProcess("Row $($i+1)/$total", "Filtering complete/empty rows")) {
                if ($i % $updateInterval -eq 0 -or $i -eq $total - 1) {
                    $percent = [int](($i+1)/$total*100)
                    Write-Progress -Activity "Filtering row $($i+1)/$total..." -PercentComplete $percent
                }
            }
            if (Test-ValidDataRow $Data[$i] $statusCol) { $Filtered += $Data[$i] }
        }
        
        if ($Filtered.Count -eq 0) { throw "All items in '$Worksheet' are complete or the sheet is empty. No report needed." }

        $Now = Get-Date

        if ($dateCol) {
            $processedData = foreach ($row in $Filtered) {
                $daysUntilDue = "N/A"
                $sortableDate = [datetime]::MaxValue
                
                $targetDate = ConvertTo-SafeDateTime $row.$dateCol
                if ($targetDate) {
                    $daysUntilDue = ($targetDate.Date - $Now.Date).Days
                    $sortableDate = $targetDate
                }
                
                $row | Add-Member -MemberType NoteProperty -Name 'Days Until Due' -Value $daysUntilDue -Force
                $row | Add-Member -MemberType NoteProperty -Name '_SortDate' -Value $sortableDate -Force
                $row
            }
            $Filtered = $processedData | Sort-Object -Property '_SortDate' | Select-Object -ExcludeProperty '_SortDate'
        }

        $escalationRecipientList = $EscalationRecipients -split ";" | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        if ($escalationRecipientList.Count -gt 0 -and $dateCol) {
            if (-not (Test-ValidEmailList -EmailArray $escalationRecipientList)) {
                Write-Log "Invalid escalation email addresses. Report generated, but escalation email not sent."
            } else {
                $Escalated = $Filtered | Where-Object { $td = ConvertTo-SafeDateTime $_.$dateCol; $null -ne $td -and ($Now.Date - $td.Date).Days -ge $EscalationDays }
                if ($Escalated.Count -gt 0) {
                    try { Send-EscalationEmail -Recipients $escalationRecipientList -EscalatedRows $Escalated -Today $Now.ToString("yyyy-MM-dd") -EscalationDays $EscalationDays }
                    catch { Write-Log "Could not send escalation email: $($_.Exception.Message)" }
                }
            }
        }

        $htmlFragment = $Filtered | ConvertTo-Html -Fragment
        $reportPaths = Publish-Report -DataToExport $Filtered -OutputFolder $OutputFolder -SourceFilePath $Path -ReportType "ReadinessReport" -HtmlBody $htmlFragment -WorksheetName $Worksheet

        $recipientList = $Recipients -split ";" | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        if ($recipientList.Count -gt 0) {
            foreach ($f in @($reportPaths.HtmlPath, $reportPaths.ExcelPath)) {
                if (-not (Test-Path $f -PathType Leaf)) { throw "Cannot send email, missing attachment: $f" }
            }
            try {
                Send-EmailSafely -To $recipientList                          -Subject "Release Readiness Review Report - $Worksheet - $(Get-Date -Format 'yyyy-MM-dd')"                          -Body "Please find attached the Release Readiness Review Report."                          -Attachments @($reportPaths.HtmlPath, $reportPaths.ExcelPath)
            } catch {
                Write-Log "Could not send main report email: $($_.Exception.Message)"
            }
        }
    } catch {
        throw "Standard report processing failed: $($_.Exception.Message)"
    }
}

function Invoke-BatchReport {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [ValidateScript({ Test-Path $_ -PathType Leaf })]
        [string]$ExcelPath,

        [Parameter(Mandatory)]
        [ValidateScript({ Test-Path $_ -PathType Container })]
        [string]$OutputFolder
    )
    Write-Log "Processing Batch Report for file: $ExcelPath"
    try {
        Confirm-ImportExcelModule
        $excelOut = Join-Path $OutputFolder "ReleaseReadiness_Requirements_Reports_$(Get-Date -Format 'yyyy-MM-dd').xlsx"
        $releaseReadinessData = @()

        $sheetInfo = (Get-ExcelSheetInfo -Path $ExcelPath -ErrorAction Stop).Name
        
        $total = $sheetInfo.Count
        $updateInterval = [Math]::Max(1, [int]($total / 100))
        for($i=0; $i -lt $total; $i++) {
            if ($pscmdlet.ShouldProcess("Sheet $($i+1)/$total", "Processing sheet for Release Readiness Review requirements")) {
                if ($i % $updateInterval -eq 0 -or $i -eq $total - 1) {
                    $percent = [int](($i+1)/$total*100)
                    Write-Progress -Activity "Processing sheet $($i+1)/$total..." -PercentComplete $percent
                }
            }

            $sheet = $sheetInfo[$i]
            
            try {
                $data = Import-Excel -Path $ExcelPath -WorksheetName $sheet -ErrorAction Stop
                if ($data -and $data.Count -gt 0) {
                    $wsCol = $data[0].PSObject.Properties.Name | Where-Object { $_ -match '(?i)(workstream|type)' } | Select-Object -First 1
                    if ($wsCol) {
                        $releaseReadinessRows = $data | Where-Object { $_.$wsCol -match "Release Readiness" }
                        if ($releaseReadinessRows) {
                            $releaseReadinessRows | ForEach-Object { $_ | Add-Member -NotePropertyName "Source_Sheet" -NotePropertyValue $sheet -Force }
                            $releaseReadinessData += $releaseReadinessRows
                        }
                    }
                }
            } catch { Write-Log "Could not process sheet '$($sheet.Name)' in batch mode. Error: $($_.Exception.Message)" }
        }
        if ($releaseReadinessData.Count -eq 0) { throw "No Release Readiness Review Requirements found in any worksheet." }
        $releaseReadinessData | Export-Excel -Path $excelOut -WorksheetName "ReleaseReadiness_Requirements" -AutoSize -AutoFilter
    } catch {
        throw "Batch report processing failed: $($_.Exception.Message)"
    }
}

function Send-EscalationEmail {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string[]]$Recipients,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [object[]]$EscalatedRows,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Today,
        
        [Parameter(Mandatory)]
        [int]$EscalationDays
    )
    
    $Headers = $EscalatedRows[0].PSObject.Properties.Name
    $bodyBuilder = New-Object System.Text.StringBuilder
    [void]$bodyBuilder.AppendLine("<h2>Escalation: Release Readiness Review Items Overdue > $($EscalationDays) Days</h2>")
    [void]$bodyBuilder.AppendLine("<table border='1' style='border-collapse:collapse;'><tr>")
    $Headers | ForEach-Object { [void]$bodyBuilder.Append("<th style='padding:5px;border:1px solid black;'>$(ConvertTo-SafeHtml $_)</th>") }
    [void]$bodyBuilder.AppendLine("</tr>")
    foreach ($item in $EscalatedRows) {
        [void]$bodyBuilder.Append("<tr>")
        $Headers | ForEach-Object { [void]$bodyBuilder.Append("<td style='padding:5px;border:1px solid black;'>$(if($item.$_) { ConvertTo-SafeHtml $item.$_ } else { '' })</td>") }
        [void]$bodyBuilder.AppendLine("</tr>")
    }
    [void]$bodyBuilder.AppendLine("</table><br><p>Please address these items urgently.</p>")
    $htmlBody = $bodyBuilder.ToString()

    Send-EmailSafely -To $Recipients          -Subject "Release Readiness Review Escalation Notification - $Today"          -HTMLBody (Get-FullHtml -fragment $htmlBody -title "Release Readiness Review Escalation Notification")
}

function Publish-Report {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [object[]]$DataToExport,

        [Parameter(Mandatory)]
        [ValidateScript({ Test-Path $_ -PathType Container })]
        [string]$OutputFolder,

        [Parameter(Mandatory)]
        [ValidateScript({ Test-Path $_ -PathType Leaf })]
        [string]$SourceFilePath,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$WorksheetName,

        [string]$ReportType = "Report",
        
        [string]$HtmlBody
    )

    try {
        $safeSheet = ($WorksheetName -replace '[\\/:*?"<>|]', '_').Trim()
        $today = Get-Date -Format "yyyy-MM-dd"
        $htmlReportPath = Join-Path $OutputFolder "ReleaseReadiness_Report_${safeSheet}_$today.html"
        $excelOutputPath = Join-Path $OutputFolder "ReleaseReadiness_Output_${safeSheet}_$today.xlsx"
        
        $archiveFolder = Join-Path $OutputFolder "Archives"
        if (-not (Test-Path $archiveFolder)) {    
            New-Item -Path $archiveFolder -ItemType Directory -Force | Out-Null    
        }
        $excelArchivePath = Join-Path $archiveFolder "ReleaseReadiness_Source_$(Get-Date -Format 'yyyyMMdd-HHmmss').xlsx"
        
        foreach ($ref in [ref]$htmlReportPath, [ref]$excelOutputPath, [ref]$excelArchivePath) {
            $p = $ref.Value
            if ($p.Length -gt 259 -and -not $p.StartsWith("\\?\\", [System.StringComparison]::Ordinal)) {
                if ($p.StartsWith("\\", [System.StringComparison]::Ordinal)) {
                    $ref.Value = "\\?\UNC\$($p.TrimStart('\'))"
                }
                else {
                    $ref.Value = "\\?\$p"
                }
            }
        }

        if ($HtmlBody) {
            $safeHtmlBody = $HtmlBody -replace '<script[^>]*>.*?</script>', '' -replace 'javascript:', ''
            $fullHtml = Get-FullHtml -fragment $safeHtmlBody -title $WorksheetName
            [System.IO.File]::WriteAllText($htmlReportPath, $fullHtml, [System.Text.Encoding]::UTF8)
        }

        if (Test-Path $excelOutputPath) {
            if (-not (Test-FileAccess -FilePath $excelOutputPath).Success) {    
                throw "Cannot overwrite locked output file: $excelOutputPath"    
            }
            Remove-Item $excelOutputPath -Force
        }

        $DataToExport | Export-Excel -Path $excelOutputPath -WorksheetName $ReportType -AutoSize -AutoFilter

        if (Test-Path $SourceFilePath) {
            try {    
                Invoke-WithRetry -ScriptBlock {
                    Copy-Item -Path $SourceFilePath -Destination $excelArchivePath -Force    
                } -MaxAttempts 3 -DelayMilliseconds 500
            } catch {    
                Write-Log "Could not archive source file after multiple attempts: $($_.Exception.Message)"    
            }
        }
        
        Write-Log "Report successfully published to $OutputFolder"
        return @{ HtmlPath = $htmlReportPath; ExcelPath = $excelOutputPath }
    } catch {
        throw "Failed to publish report: $($_.Exception.Message)"
    }
}

#endregion

$SettingsFolder     = Join-Path $env:LOCALAPPDATA 'ReleaseReadinessReportGenerator'
if (-not (Test-Path $SettingsFolder)) { New-Item -Path $SettingsFolder -ItemType Directory -Force | Out-Null }
$settingsFilePath = Join-Path $SettingsFolder 'settings.json'
$DefaultOutputFolder = [System.IO.Path]::Combine($env:USERPROFILE, "Documents", "ReleaseReadiness_Reports")
$SpecialScenarioSheets = @("Scenario 1 - Checklist", "Scenario 2 - Checklist")
$ColumnAliases = @{
    Status      = @('Status', 'State', 'Progress')
    'Target Date' = @('Target Date', 'Due Date', 'Deadline', 'ETA')
    Requirement = @('Requirement', 'Component Task', 'Description', 'Release Readiness Requirements')
}
$ColumnAliases['Requirement Status'] = @('Requirement Status', 'Status', 'State', 'Progress')
$ColumnAliases['Assigned to']        = @('Assigned to', 'Owner', 'Responsible')

$olMailItem = 0
$olHTML     = 2

#region Form and Layout Initialization
$form = New-Object System.Windows.Forms.Form
$form.Text = "Release Readiness Review Report Generator"
$form.Size = New-Object System.Drawing.Size(800, 900)
$form.StartPosition = "CenterScreen"
$form.AutoScaleMode = 'Dpi'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.MinimizeBox = $true

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 200

$script:pwshInstance = $null
$script:asyncResult = $null
$script:OperationStartTime = $null

$tableLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$tableLayoutPanel.Dock = "Fill"
$tableLayoutPanel.Padding = New-Object System.Windows.Forms.Padding(10)
$form.Controls.Add($tableLayoutPanel)

$tableLayoutPanel.ColumnCount = 2
[void]$tableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 300)))
[void]$tableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))

$rowHeights = @(60, 50, 40, 40, 40, 40, 40, 50, 20, 70, 60, 50, 40, 40, 40)
$tableLayoutPanel.RowCount = $rowHeights.Count
foreach ($height in $rowHeights) {
    [void]$tableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, $height)))
}
#endregion

#region GUI Controls
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Release Readiness Review Report Generator"
$titleLabel.Font = New-Object System.Drawing.Font("Calibri", 20, [System.Drawing.FontStyle]::Bold)
$titleLabel.Anchor = "None"
$titleLabel.TextAlign = "MiddleCenter"
$titleLabel.Dock = "Fill"

$selectButton = New-Object System.Windows.Forms.Button
$selectButton.Text = "Select Release Readiness Excel File"
$selectButton.Font = New-Object System.Drawing.Font("Segoe UI", 11)
$selectButton.Anchor = "None"
$selectButton.Size = New-Object System.Drawing.Size(280, 40)

$labelPath = New-Object System.Windows.Forms.Label
$labelPath.Text = "No Release Readiness Excel file selected. Please select one."
$labelPath.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$labelPath.Dock = "Fill"
$labelPath.TextAlign = "MiddleCenter"

$worksheetLabel = New-Object System.Windows.Forms.Label
$worksheetLabel.Text = "Select Worksheet:"
$worksheetLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$worksheetLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Right
$worksheetLabel.AutoSize = $true

$worksheetComboBox = New-Object System.Windows.Forms.ComboBox
$worksheetComboBox.DropDownStyle = 'DropDownList'
$worksheetComboBox.Dock = "Fill"
$worksheetComboBox.Anchor = [System.Windows.Forms.AnchorStyles]::None

$emailLabel = New-Object System.Windows.Forms.Label
$emailLabel.Text = "Email Recipients (semicolon-separated):"
$emailLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$emailLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Right
$emailLabel.AutoSize = $true

$emailTextBox = New-Object System.Windows.Forms.TextBox
$emailTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$emailTextBox.Dock = "Fill"
$emailTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::None

$escalationLabel = New-Object System.Windows.Forms.Label
$escalationLabel.Text = "Escalation Threshold (Days Overdue):"
$escalationLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$escalationLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Right
$escalationLabel.AutoSize = $true

$escalationNumeric = New-Object System.Windows.Forms.NumericUpDown
$escalationNumeric.Minimum = 1
$escalationNumeric.Maximum = 365
$escalationNumeric.Value = 7
$escalationNumeric.Anchor = "Left"

$escalationEmailsLabel = New-Object System.Windows.Forms.Label
$escalationEmailsLabel.Text = "Escalation Emails (semicolon-separated):"
$escalationEmailsLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$escalationEmailsLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Right
$escalationEmailsLabel.AutoSize = $true

$escalationEmailsTextBox = New-Object System.Windows.Forms.TextBox
$escalationEmailsTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$escalationEmailsTextBox.Dock = "Fill"
$escalationEmailsTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::None

$outputLabel = New-Object System.Windows.Forms.Label
$outputLabel.Text = "Output Folder:"
$outputLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$outputLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Right
$outputLabel.AutoSize = $true

$outputPanel = New-Object System.Windows.Forms.Panel
$outputPanel.Dock = "Fill"

$outputPathBox = New-Object System.Windows.Forms.TextBox
$outputPathBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$outputPathBox.Text = $DefaultOutputFolder
$outputPathBox.Dock = "Fill"

$heightValue = if ($null -ne $outputPathBox) { $outputPathBox.Height + 4 } else { 28 }
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Browse..."
$browseButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$browseButton.Size = New-Object System.Drawing.Size(90, $heightValue)
$browseButton.Dock = "Right"

$runButton = New-Object System.Windows.Forms.Button
$runButton.Text = "Run Readiness Report"
$runButton.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
$runButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$runButton.ForeColor = [System.Drawing.Color]::White
$runButton.Anchor = "None"
$runButton.Size = New-Object System.Drawing.Size(250, 54)

$batchButton = New-Object System.Windows.Forms.Button
$batchButton.Text = "Generate Release Readiness Requirements Report (Batching)"
$batchButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$batchButton.Anchor = "None"
$batchButton.Size = New-Object System.Drawing.Size(410, 39)

$saveButton = New-Object System.Windows.Forms.Button
$saveButton.Text = "Save Settings"
$saveButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$saveButton.Anchor = "None"
$saveButton.Size = New-Object System.Drawing.Size(150, 32)

$processingLabel = New-Object System.Windows.Forms.Label
$processingLabel.Text = "Processing, please wait..."
$processingLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Italic)
$processingLabel.ForeColor = [System.Drawing.Color]::DarkBlue
$processingLabel.Visible = $false
$processingLabel.Anchor = "None"
$processingLabel.TextAlign = "MiddleCenter"
$processingLabel.Dock = "Fill"

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Cancel"
$cancelButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$cancelButton.Anchor = "None"
$cancelButton.Size = New-Object System.Drawing.Size(150, 32)
$cancelButton.Visible = $false

Add-Control -control $titleLabel -row 0 -col 0 -colSpan 2
Add-Control -control $selectButton -row 1 -col 0 -colSpan 2
Add-Control -control $labelPath -row 2 -col 0 -colSpan 2
Add-control -control $worksheetLabel -row 3 -col 0
Add-Control -control $worksheetComboBox -row 3 -col 1
Add-Control -control $emailLabel -row 4 -col 0
Add-Control -control $emailTextBox -row 4 -col 1
Add-Control -control $escalationLabel -row 5 -col 0
Add-Control -control $escalationNumeric -row 5 -col 1
Add-Control -control $escalationEmailsLabel -row 6 -col 0
Add-Control -control $escalationEmailsTextBox -row 6 -col 1
Add-Control -control $outputLabel -row 7 -col 0
Add-Control -control $outputPanel -row 7 -col 1
[void]$outputPanel.Controls.Add($outputPathBox)
[void]$outputPanel.Controls.Add($browseButton)
Add-Control -control $runButton -row 10 -col 0 -colSpan 2
Add-Control -control $batchButton -row 11 -col 0 -colSpan 2
Add-Control -control $saveButton -row 12 -col 0 -colSpan 2
Add-Control -control $processingLabel -row 13 -col 0 -colSpan 2
Add-Control -control $cancelButton -row 14 -col 0 -colSpan 2
#endregion

#region Event Handlers

$timer.Add_Tick({
    if ($script:OperationStartTime) {
        $elapsed = (Get-Date) - $script:OperationStartTime
        if ($elapsed.TotalMinutes -gt 10) {
            $timer.Stop()
            if ($script:pwshInstance -and -not $script:asyncResult.IsCompleted) {
                $script:pwshInstance.Stop()
            }
            Show-UIMessage -Message "Operation timed out after 10 minutes. The process has been cancelled." -Title "Timeout" -Icon Warning
            
            if ($script:pwshInstance) {
                try {
                    $script:pwshInstance.Runspace.Dispose()
                    $script:pwshInstance.Dispose()
                } catch {
                    Write-Log "Error during timeout cleanup: $($_.Exception.Message)"
                }
            }
            $script:pwshInstance = $null
            $script:asyncResult = $null
            $script:OperationStartTime = $null
            
            $form.Enabled = $true
            $processingLabel.Visible = $false
            $cancelButton.Visible = $false
            $cancelButton.Enabled = $true
            $cancelButton.Text = "Cancel"
            return
        }
    }
    
    if ($script:pwshInstance -and $script:pwshInstance.Streams.Progress.Count -gt 0) {
        $progress = $script:pwshInstance.Streams.Progress.Read(1)
        if ($progress) {
            $processingLabel.Text = $progress.Activity
        }
    }

    if ($script:asyncResult -and $script:asyncResult.IsCompleted) {
        $timer.Stop()
        try {
            [void]$script:pwshInstance.EndInvoke($script:asyncResult)

            if ($script:pwshInstance.HadErrors) {
                $errors = $script:pwshInstance.Streams.Error -join "`n"
                Show-UIMessage -Message ("Report generation failed:`n" + $errors) -Title "Report Generation Error" -Icon Error
            } else {
                Show-UIMessage -Message "Report completed successfully!" -Title "Success" -Icon Information
            }
        } catch {
            $errMsg = $_.Exception.ToString()
            Show-UIMessage -Message ("An unexpected error occurred:`n" + $errMsg) -Title "Worker Error" -Icon Error
        } finally {
            if ($script:pwshInstance) {
                try {
                    $script:pwshInstance.Runspace.Dispose()
                    $script:pwshInstance.Dispose()
                } catch {
                    Write-Log "Error during final cleanup: $($_.Exception.Message)"
                }
            }
            $script:pwshInstance = $null
            $script:asyncResult = $null
            $script:OperationStartTime = $null

            $form.Enabled = $true
            $processingLabel.Visible = $false
            $cancelButton.Visible = $false
            $cancelButton.Enabled = $true
            $cancelButton.Text = "Cancel"
        }
    }
})

function Start-AsyncReport {
    param(
        [Parameter(Mandatory)]
        [string]$CommandName,
        [Parameter(Mandatory)]
        [hashtable]$CommandParameters
    )
    
    try {
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        foreach ($kvp in $script:SessionStateInfo.Functions.GetEnumerator()) {
            $iss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new($kvp.Name, $kvp.Value)) | Out-Null
        }
        foreach ($kvp in $script:SessionStateInfo.Aliases.GetEnumerator()) {
            $iss.Commands.Add([System.Management.Automation.Runspaces.SessionStateAliasEntry]::new($kvp.Name, $kvp.Value)) | Out-Null
        }

        $runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($iss)
        $runspace.ApartmentState = 'STA'
        $runspace.ThreadOptions = 'ReuseThread'
        $runspace.Open()
        
        $script:pwshInstance = [powershell]::Create()
        $script:pwshInstance.Runspace = $runspace
        
        $script:pwshInstance.AddCommand($CommandName).AddParameters($CommandParameters) | Out-Null

        $script:asyncResult = $script:pwshInstance.BeginInvoke()
        
        $script:OperationStartTime = Get-Date

        $form.Enabled = $false
        $processingLabel.Visible = $true
        $processingLabel.Text = "Processing, please wait..."
        $cancelButton.Visible = $true
        $timer.Start()
    } catch {
        if ($script:pwshInstance) {
            try {
                $script:pwshInstance.Runspace.Dispose()
                $script:pwshInstance.Dispose()
            } catch {
                Write-Log "Error during cleanup: $($_.Exception.Message)"
            }
            $script:pwshInstance = $null
            $script:asyncResult = $null
        }
        throw
    }
}

$selectButton.Add_Click({
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    try {
        $fileDialog.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls"
        $fileDialog.Title = "Select Release Readiness Excel File"
        if ($fileDialog.ShowDialog($form) -eq "OK") {
            $selectedFile = $fileDialog.FileName

            if (-not (Test-Path $selectedFile -PathType Leaf)) {
                throw "Excel file not found: $selectedFile"
            }
            $lock = Test-FileAccess -FilePath $selectedFile
            if (-not $lock.Success) {
                throw "Cannot open locked Excel file: $($lock.Error)"
            }

            $script:SelectedFilePath = $selectedFile
            $labelPath.Text = "Selected File: " + [System.IO.Path]::GetFileName($selectedFile)
            Write-Log "User selected file: $($script:SelectedFilePath)"
            
            $excelApp = $null
            $wb = $null
            $sheetList = $null
            try {
                $excelApp = New-Object -ComObject Excel.Application
                $excelApp.Visible = $false
                $wb = $excelApp.Workbooks.Open($selectedFile, 0, $true)
                $sheetList = $wb.Sheets | ForEach-Object { $_.Name }
                
                foreach ($ws in @($wb.Sheets)) {
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
                }
                
                $wb.Close($false)
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
                $wb = $null
            }
            catch {
                Write-Log "COM-sheet read failed: $($_.Exception.Message). Falling back to Get-ExcelSheetInfo."
                try {
                    Confirm-ImportExcelModule
                    $sheetList = (Get-ExcelSheetInfo -Path $selectedFile -ErrorAction Stop).Name
                }
                catch {
                    $errMsg = "Failed to read worksheets from Excel file using any method: $($_.Exception.ToString())"
                    Write-Log $errMsg
                    Show-UIMessage -Message $errMsg -Title "Excel File Error" -Icon Error
                }
            }
            finally {
                if ($null -ne $wb) {
                    try {
                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
                    } catch {}
                }
                if ($null -ne $excelApp) {
                    try {
                        $excelApp.Quit()
                    } catch {
                        Write-Log "Warning: Could not call Quit on Excel. Error: $($_.Exception.Message)"
                    } finally {
                        try {
                            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp) | Out-Null
                        } catch {
                            Write-Log "Warning: Could not release Excel COM object"
                        }
                        $excelApp = $null
                        [System.GC]::Collect()
                        [System.GC]::WaitForPendingFinalizers()
                    }
                }
            }
            
            Test-ControlInitialization -ControlName 'worksheetComboBox' -Control $worksheetComboBox
            $worksheetComboBox.Items.Clear()
            if ($sheetList) {
                $sheetList | ForEach-Object { [void]$worksheetComboBox.Items.Add($_) }
                if ($worksheetComboBox.Items.Count -gt 0) {
                    $worksheetComboBox.SelectedIndex = 0
                }
            }
        }
    }
    finally {
        if ($null -ne $fileDialog) { $fileDialog.Dispose() }
    }
})

$browseButton.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    try {
        $folderDialog.Description = "Select Output Folder for Reports"
        $currentPath = Get-SafeControlValue -ControlName 'outputPathBox' -Control $outputPathBox -Property 'Text' -DefaultValue $DefaultOutputFolder
        $folderDialog.SelectedPath = $currentPath
        if ($folderDialog.ShowDialog($form) -eq "OK") {
            Set-SafeControlValue -ControlName 'outputPathBox' -Control $outputPathBox -Value $folderDialog.SelectedPath -Property 'Text'
            Write-Log "User selected output folder: $($folderDialog.SelectedPath)"
        }
    }
    finally {
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
        $requiredControls = @{ 'worksheetComboBox' = $worksheetComboBox; 'outputPathBox' = $outputPathBox; 'emailTextBox' = $emailTextBox; 'escalationNumeric' = $escalationNumeric; 'escalationEmailsTextBox' = $escalationEmailsTextBox }
        foreach ($controlPair in $requiredControls.GetEnumerator()) { Test-ControlInitialization -ControlName $controlPair.Key -Control $controlPair.Value }
        
        $path = $script:SelectedFilePath
        if (-not $path) { throw "Please select an Excel file first." }
        
        $lock = Test-FileAccess -FilePath $path
        if (-not $lock.Success) { throw "The selected file '$([System.IO.Path]::GetFileName($path))' is locked or inaccessible. Error: $($lock.Error)" }
        
        $sheet = $worksheetComboBox.SelectedItem
        if (-not $sheet) { throw "Please select a worksheet." }
        
        $outputFolder = Get-SafeControlValue -ControlName 'outputPathBox' -Control $outputPathBox -Property 'Text' -DefaultValue $DefaultOutputFolder
        if (-not (Test-ValidPath $outputFolder)) {
            try { New-Item -Path $outputFolder -ItemType Directory -Force | Out-Null } catch { throw "Output folder is not valid and could not be created: $($_.Exception.Message)" }
        }

        $params = @{
            Path                 = $path
            Worksheet            = [string]$sheet
            OutputFolder         = [string]$outputFolder
            Recipients           = [string](Get-SafeControlValue -ControlName 'emailTextBox' -Control $emailTextBox -Property 'Text')
            EscalationDays       = [int]$escalationNumeric.Value
            EscalationRecipients = [string](Get-SafeControlValue -ControlName 'escalationEmailsTextBox' -Control $escalationEmailsTextBox -Property 'Text')
        }
        
        $normSpecial = $SpecialScenarioSheets | ForEach-Object { Convert-Name $_ }
        $commandName = if (Convert-Name $sheet -in $normSpecial) { 'Invoke-SpecialReport' } else { 'Invoke-StandardReport' }
        
        Start-AsyncReport -CommandName $commandName -CommandParameters $params

    } catch {
        $errorMsg = "Report generation failed: $($_.Exception.ToString())"
        Write-Log "Pre-flight check failed for Run Report: $($_.Exception.Message)"
        Show-UIMessage -Message $errorMsg -Title "Report Generation Error" -Icon Error
    }
})

$batchButton.Add_Click({
    try {
        Write-Log "'Batch Report' button clicked."
        Test-ControlInitialization -ControlName 'outputPathBox' -Control $outputPathBox
        $excelPath = $script:SelectedFilePath
        if (-not $excelPath) { throw "Please select an Excel file first." }
        
        $outputFolder = Get-SafeControlValue -ControlName 'outputPathBox' -Control $outputPathBox -Property 'Text' -DefaultValue $DefaultOutputFolder
        if (-not (Test-ValidPath $outputFolder)) {
            try { New-Item -Path $outputFolder -ItemType Directory -Force | Out-Null } catch { throw "Invalid or inaccessible output folder: $($_.Exception.Message)" }
        }

        $params = @{
            ExcelPath    = $excelPath
            OutputFolder = [string]$outputFolder
        }
        
        Start-AsyncReport -CommandName 'Invoke-BatchReport' -CommandParameters $params

    } catch {
        $errorMsg = "Batch report failed: $($_.Exception.ToString())"
        Write-Log "Pre-flight check failed for Batch Report: $($_.Exception.Message)"
        Show-UIMessage -Message $errorMsg -Title "Batch Report Error" -Icon Error
    }
})
#endregion

#region Application Startup and Shutdown
try {
    $commandLineArgs = [Environment]::GetCommandLineArgs()
    if ($commandLineArgs -contains '-Verbose' -or $commandLineArgs -contains '/Verbose') {
        $global:DebugPreference = 'Continue'
        $global:VerbosePreference = 'Continue'
    }

    Write-Log "--- Application starting ---"
    
    if (-not $script:IsCompiledEXE -and $PSVersionTable.PSEdition -eq 'Core') {
        Show-UIMessage -Message "This tool requires Windows PowerShell 5.1 for full COM support." -Title "Initialization Error" -Icon ([System.Windows.Forms.MessageBoxIcon]::Error)
        exit 1
    }
    
    try {
        if (-not (Test-Path $DefaultOutputFolder)) {
            New-Item -Path $DefaultOutputFolder -ItemType Directory -Force | Out-Null
        }
    } catch {
        Show-UIMessage -Message "Cannot create or access output folder:`n${DefaultOutputFolder}`nError: $($_.Exception.Message)" `
                         -Title "Permission Error" -Icon ([System.Windows.Forms.MessageBoxIcon]::Error)
        exit 1
    }

    $script:SessionStateInfo = @{
        Functions = @{};
        Aliases = @{}
    }
    
    $functionsToExport = @(
        'Write-Log', 'Show-UIMessage', 'Confirm-ImportExcelModule',
        'Test-FileAccess', 'Test-ValidPath', 'Test-ValidEmailList',
        'Test-ValidDataRow', 'ConvertTo-SafeHtml', 'ConvertTo-SafeDateTime',
        'Convert-Name', 'Get-FullHtml', 'Test-OutlookAvailability',
        'Send-EmailSafely', 'Test-ControlInitialization', 'Get-SafeControlValue',
        'Set-SafeControlValue', 'Save-Setting', 'Import-Setting', 'Add-Control',
        'Invoke-WithRetry', 'Invoke-SpecialReport', 'Invoke-StandardReport',
        'Invoke-BatchReport', 'Send-EscalationEmail', 'Publish-Report', 'Test-FormExists'
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

    $coreControls = @{ 'form' = $form; 'emailTextBox' = $emailTextBox; 'escalationNumeric' = $escalationNumeric; 'escalationEmailsTextBox' = $escalationEmailsTextBox; 'outputPathBox' = $outputPathBox; 'worksheetComboBox' = $worksheetComboBox }
    Write-Log "Validating GUI control initialization..."
    foreach ($controlPair in $coreControls.GetEnumerator()) { [void](Test-ControlInitialization -ControlName $controlPair.Key -Control $controlPair.Value) }
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
    }))

    [void][System.Windows.Forms.Application]::Run($form)
}
catch {
    $full = $_.Exception.ToString()
    Write-Log "Failed to initialize application: $full"
    [System.Windows.Forms.MessageBox]::Show(
        "Initialization failed:`n$full",
        "Initialization Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit 1
}
finally {
    if ($timer) { $timer.Dispose() }
    if ($form) { $form.Dispose() }
    if ($script:pwshInstance) {
        try {
            $script:pwshInstance.Runspace.Dispose()
            $script:pwshInstance.Dispose()
        } catch {
            Write-Log "Error during final cleanup: $($_.Exception.Message)"
        }
    }
    Write-Log "--- Application terminated ---`n"
}
#endregion
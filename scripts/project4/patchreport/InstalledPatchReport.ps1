param(
    $AuthUser = "administrator",
    $AuthPassword = "Wigtra@2020",
    $ToEmail = "gpk.pandey.neeraj@gmail.com",
    $FromEmail = "no_reply@notification.com"
)

$SMTP_USER = "abc@gmail.com"
$SMTP_PASS = "abc"

Function Start-Log {
    param(
        [string]$LogFileName
    )
    $Location = (Join-Path $PSScriptRoot $LogFileName)
    New-Item -ItemType file $Location -Force | Out-Null
    $Global:LogFileLocation = $Location
}

Function Log-Info {
    param(
        [string]$Message
    )
    
    $logEntry = [pscustomobject]@{
        log_level = "INFO"
        message   = $Message
    }
    $logEntry | Export-Csv -Path $Global:LogFileLocation -Encoding ASCII -Append -NoTypeInformation -Force 
}

Function Log-Error {
    param(
        [string]$Message
    )
    $logEntry = [pscustomobject]@{
        log_level = "ERROR"
        message   = $Message
    }
    $logEntry | Export-Csv -Path $Global:LogFileLocation -Encoding ASCII -Append -NoTypeInformation -Force
}

Function Archive-Report {
    param(
        $FileName
    )
    $ArchiveDir = Join-Path $PSScriptRoot "Archive"
    [System.IO.Directory]::CreateDirectory($ArchiveDir) | Out-Null
    if (Test-Path $FileName) {
        Move-Item $FileName -Destination $ArchiveDir -Force
    }
}

$GetInstalledPatches = {
    Get-HotFix -Id $patches -ErrorAction SilentlyContinue
}

function Send-StageStatusEmail {
    param (
        $FromEmail,
        $ToEmail,
        $Subject,
        $ReportPath
    )
    
    $logs = Import-Csv $Global:LogFileLocation
    $Password = $SMTP_PASS | ConvertTo-SecureString -AsPlainText -Force
    $AuthUser = $SMTP_USER
    $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AuthUser, $Password
    $EmailParams = @{
        from       = $FromEmail
        to         = $ToEmail
        subject    = $Subject
        smtpserver = "smtp-relay.brevo.com"
        port       = 587
        UseSsl     = $true
        credential = $Credential
        BodyAsHtml = $true
    }

    $MailMessage = '<!DOCTYPE html><html><head></head><body>'
    if (-not (test-path $ReportPath)) {
        $MailMessage += '<p>Hello Team,<br><br>Your initiated, logged in users report, has been completed without final report. Please contact your administrator.</br></br></p>'
    }
    else {
        $MailMessage += '<p>Hello Team,<br><br>Your initiated, logged in users repor, has been completed. Please check the logs below and take appropriate action.</br></br></p>'
        $EmailParams['Attachments'] = $ReportPath
    }

    $MailMessage += '<h3>Logs</h3><table style="font-family: arial, sans-serif;border-collapse: collapse;width: 50%;">'
    $MailMessage += '<tr><th style="border: 1px solid #dddddd;text-align: left;padding: 8px;word-wrap:break-word;">Type</th>'
    $MailMessage += '<th style="border: 1px solid #dddddd;text-align: left;padding: 8px;word-wrap:break-word;">Message</th></tr>'

    for ($i = 0; $i -lt $logs.Length; $i++) {
        if ($logs[$i].log_level -eq 'PROGRESS') { continue }
        $bgcolor = if ($i % 2 -eq 0) { "#dddddd" } else { "#FFFFFF" }
        $msgcolor = if ($logs[$i].log_level -eq 'ERROR') { "red" } else { "black" }
        $MailMessage += '<tr style="background-color:{0}"><td style="border: 1px solid #dddddd;text-align: left;padding: 8px;word-wrap:break-word; color:{1};">{2}</td>' -f $bgcolor, $msgcolor, $logs[$i].log_level
        $MailMessage += '<td style="border: 1px solid #dddddd;text-align: left;padding: 8px;word-wrap:break-word;">{0}</td></tr>' -f $logs[$i].message
    }
    $MailMessage += "</table></body></html></table><br><br>Regards,<br>Windows Migration Team"
    
    $EmailParams['body'] = $MailMessage
    
    Send-MailMessage @EmailParams
}

$InstalledPatchReportFile = Join-Path $PSScriptRoot "InstalledPatchReport.xlsx"
Start-Log -LogFileName "Log_InstalledPatch.csv"
Archive-Report $InstalledPatchReportFile
try {

    [array]$Computers = Get-Content (Join-Path $PSScriptRoot "hosts.txt") | Select-Object -Unique
    [array]$Patches = Get-Content (Join-Path $PSScriptRoot "patches.txt") | Select-Object -Unique

    $SecurePassword = $AuthPassword | ConvertTo-SecureString -AsPlainText -Force
    $Credential = New-Object System.Management.Automation.PSCredential($AuthUser, $SecurePassword)
    $InstalledPatchReport = @()
    foreach ($computer in $Computers) {
        try {
            $computer = $computer.trim()
            if ([bool](Test-WSMan -ComputerName $Computer -ErrorAction SilentlyContinue)) {
                Log-Info -Message "Querying host '$computer' for installed patches." 
                $InvokeGetInstalledPatches = Invoke-Command -ComputerName $Computer -Credential $Credential -ScriptBlock $GetInstalledPatches -ArgumentList $Patches
                if (-not $InvokeGetInstalledPatches) {
                    foreach ($Patch in $Patches) {
                        $InstalledPatchReport += New-Object psobject -Property @{
                            ComputerName = $computer
                            Patch        = $Patch
                            Status       = "NO"            
                        }
                    }
                    Log-Info -Message "host '$computer' returned an empty patch list"
                }
                else {
                    foreach ($Patch in $Patches) {
                        $hotfixids = $InvokeGetInstalledPatches | Select-Object -expand HotFixID
                        if ($patch -in $hotfixids) {
                            $status = "YES"
                        }
                        else {
                            $status = "NO"
                        }
                        $InstalledPatchReport += New-Object psobject -Property @{
                            ComputerName = $computer
                            Patch        = $Patch
                            Status       = $status
                        }
                    }
                }
            }
            else {
                Log-Error -Message "Can not connect to $Computer. Make sure the hostname is valid and PS-Remoting is allowed on the target machine."
            }
        }
        catch {
            Log-Error -Message "Failed to query user profile list from host."
            Log-Error -Message $("ErrorRecord : {0}, CommandName: {1}, Message: {2}" -f $_.Exception.ErrorRecord, $_.Exception.CommandName, $_.Exception.Message)
        }
    }
    if ($InstalledPatchReport) {
        Log-Info -Message "Generating consolidated report."
        $InstalledPatchReport | Export-Excel $InstalledPatchReportFile -WorksheetName "InstalledPatchReport"
    }
}
catch {
    Log-Error -Message $("ErrorRecord : {0}, CommandName: {1}, Message: {2}" -f $_.Exception.ErrorRecord, $_.Exception.CommandName, $_.Exception.Message)
}
finally {
    Send-StageStatusEmail -FromEmail $FromEmail -ToEmail $($ToEmail -split ",") -Subject "Installed Patch Report." -ReportPath $InstalledPatchReportFile
}

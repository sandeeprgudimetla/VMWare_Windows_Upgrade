param(
    #GuestParams
    $Computer = "REPLACE_ME",
    $AuthUser = "REPLACE_ME",
    $AuthPassword = "REPLACE_ME",
    $Workspace = "C:\Temp",
    $LocalUserAccount = "REPLACE_ME",
    $ToEmail = "",
    $FromEmail = ""
)

$SecurePassword = $AuthPassword | ConvertTo-SecureString -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential($AuthUser, $SecurePassword)
Function Start-Log{
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

Function Log-Progress{
    param(
        [Parameter(ValueFromPipeline=$true)]
        [string]$Percentage
    ) 
    $logEntry = [pscustomobject]@{
        log_level = "PROGRESS"
        message   = $Percentage
    }
    
    $logEntry | Export-Csv -Path $Global:LogFileLocation -Encoding ASCII -Append -NoTypeInformation -Force
}

Start-Log -LogFileName "stage4.csv"


$DeleteNewLocalUser = {
    try{
        $localUserSource = [ADSI]"WinNT://$env:COMPUTERNAME"
        $localUserSource.delete('user', $using:LocalUserAccount)
    }
    catch {
        $_.Exception.Message
    }

}

$ConfirmDeleteNewLocalUser = {
    $localUsers = [ADSI]"WinNT://$env:COMPUTERNAME"
    [bool]($localUsers.Children | Where-Object { $_.SchemaClassName -eq 'user' -and $_.Name -eq $using:LocalUserAccount })
}


$WorkspaceCleanUp = {
    try {
        Remove-Item -Recurse -Force $using:Workspace
    }
    catch {
        $_.Exception.Message
    }
}

$ConfirmWorkspaceCleanup = {
    [bool](Test-Path $using:Workspace -ErrorAction SilentlyContinue)
}

function Send-StageStatusEmail{
    param (
        $FromEmail,
        $ToEmail,
        $Subject
    )

    $logs = Import-Csv $Global:LogFileLocation
    $MailMessage = '<!DOCTYPE html><html><head></head><body>'
    $MailMessage += '<p>Hello Team,</br></br>Your initiated stage-4 for Windows inplace upgrade is Completed. Please check the logs below and take appropriate action.</br></br></p>'
    $MailMessage += '<h3>Logs</h3><table style="font-family: arial, sans-serif;border-collapse: collapse;width: 50%;">'
    $MailMessage += '<tr><th style="border: 1px solid #dddddd;text-align: left;padding: 8px;word-wrap:break-word;">Type</th>'
    $MailMessage += '<th style="border: 1px solid #dddddd;text-align: left;padding: 8px;word-wrap:break-word;">Message</th></tr>'

    for($i=0;$i -lt $logs.Length; $i++){
        if($logs[$i].log_level -eq 'PROGRESS'){continue}
        $bgcolor = if ($i % 2 -eq 0) {"#dddddd"} else {"#FFFFFF"}
        $msgcolor = if($logs[$i].log_level -eq 'ERROR') {"red"} else {"black"}
        $MailMessage += '<tr style="background-color:{0}"><td style="border: 1px solid #dddddd;text-align: left;padding: 8px;word-wrap:break-word; color:{1};">{2}</td>' -f $bgcolor, $msgcolor, $logs[$i].log_level
        $MailMessage += '<td style="border: 1px solid #dddddd;text-align: left;padding: 8px;word-wrap:break-word;">{0}</td></tr>' -f $logs[$i].message
    }
    $MailMessage +="</table></body></html></table><br><br>Regards,<br>Windows Migration Team"
    $Password = "WbcP6wz0ktZJMpUx" | ConvertTo-SecureString -AsPlainText -Force
    $AuthUser = "anandaraaj.parthiban@gmail.com"
    $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AuthUser, $Password

    $EmailParams = @{
        from = $FromEmail
        to = $ToEmail
        subject = $Subject
        body = $MailMessage
        smtpserver = "smtp-relay.brevo.com"
        port = 587
        UseSsl = $true
        credential = $Credential
        BodyAsHtml = $true
    }
    Send-MailMessage @EmailParams
}


[array]$Computers = $Computer -split "," | Select-Object -Unique
if ($Computers.Length -gt 5) {
    Log-Error -Message "Only 5 servers are allowed for upgrade."
    exit
}

foreach ($Computer in $computers)
{
    $computer = $Computer.trim()
    Log-Info -Message "Performing cleanup for computer '$Computer'."
    if ([bool](Test-WSMan -ComputerName $Computer -ErrorAction SilentlyContinue)) {
        try {
            Log-Info -Message "Cleaning up workspace."
            $InvokeWorkspaceCleanUp = Invoke-Command -ComputerName $Computer -Credential $Credential -ScriptBlock $WorkspaceCleanUp
            $InvokeConfirmWorkspaceCleanup = Invoke-Command -ComputerName $Computer -Credential $Credential -ScriptBlock $ConfirmWorkspaceCleanup

            if($InvokeConfirmWorkspaceCleanup -eq $true){
                Log-Error -Message "Workspace clean up failed."
                Log-Error -Message $InvokeWorkspaceCleanUp
            }
            Log-Info "Workspace successfully cleaned up."
            Log-Info "Removing local user."

            $InvokeDeleteNewLocalUser = Invoke-Command -ComputerName $Computer -Credential $Credential -ScriptBlock $DeleteNewLocalUser
            $InvokeConfirmDeleteNewLocalUser = Invoke-Command -ComputerName $Computer -Credential $Credential -ScriptBlock $ConfirmDeleteNewLocalUser

            if($InvokeConfirmDeleteNewLocalUser -eq $true){
                Log-Error "Failed to remove local user."
                Log-Error -Message $InvokeDeleteNewLocalUser
            }
            Log-Info "Successfully removed local user."

        }
        catch {
            $_.Exception | Select-Object ErrorRecord, CommandName, Message | Format-List * -Force
        }
    }
    else {
        Log-Error -Message "Can not connect to $Computer. Make sure the hostname is valid and PS-Remoting is allowed on the target machine."
    }
    Log-Info "Sending email."
    Send-StageStatusEmail -FromEmail $FromEmail -ToEmail $($ToEmail -split ",") -Subject "Windows Upgrade Status: Stage 4"
}
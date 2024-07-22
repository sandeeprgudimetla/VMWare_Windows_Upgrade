param(
    #GuestParams
    $Computer = "REPLACE_ME",
    $AuthUser = "REPLACE_ME",
    $AuthPassword = "REPLACE_ME",
    $ISOLocation = "REPLACE_ME", #E.g "\\192.168.31.90\share"
    $ISOFullName = "REPLACE_ME", #E.g. "SRV2016.STD.ENU.MAY2022.iso"
    $Workspace = "C:\Temp"
)

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

$ErrorActionPreference = "Stop"
$SecurePassword = $AuthPassword | ConvertTo-SecureString -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential($AuthUser, $SecurePassword)

Start-Log -LogFileName "stage2.csv"

$CopyISOWithProgress = {
    try {
        #Create location if does not exist.
        [System.IO.Directory]::CreateDirectory($using:Workspace) | Out-Null
        #Check if drive is already mounted
        $Drive = Get-PSDrive | Where-Object { $_.Root -eq $using:ISOLocation }
        if (-not $Drive) {
            New-PSDrive -Name "Z" -PSProvider FileSystem -Root $using:ISOLocation -Credential $using:Credential -Persist | Out-Null
        }

        #Check if file is already available
        $ISO = [bool](Get-ChildItem $(Join-Path -Path $using:Workspace -ChildPath $using:ISOFullName) -ErrorAction SilentlyContinue)
        if ($ISO -eq $false) { 
            $From = Join-Path -Path "z:" -ChildPath $using:ISOFullName
            $To = Join-Path -Path $using:Workspace -ChildPath $using:ISOFullName
            $ffile = [io.file]::OpenRead($From)
            $tofile = [io.file]::OpenWrite($To)
            #Write-Progress -Activity "Copying file" -status "$From -> $To" -PercentComplete 0
            "[PROGRESS]{0}" -f 0
            try {
                [byte[]]$buff = new-object byte[] 4096
                [long]$total = [int]$count = 0

                do {
                    $count = $ffile.Read($buff, 0, $buff.Length)
                    $tofile.Write($buff, 0, $count)
                    $total += $count
                    if ($total % 1mb -eq 0) {
                        #Write-Progress -Activity "Copying file" -status "$From -> $To" `
                        #    -PercentComplete ([long]($total * 100 / $ffile.Length))

                        #"[PROGRESS]{0}" -f ([long]($total * 100 / $ffile.Length))
                        $progress = ([long]($total * 100 / $ffile.Length))
                        [pscustomobject]@{
                            log_level = "PROGRESS"
                            message   = $progress
                        }
                    }
                } while ($count -gt 0)
            }
            finally {
                $ffile.Dispose()
                $tofile.Dispose()
                #Write-Progress -Activity "Copying file" -Status "Ready" -Completed
                [pscustomobject]@{
                    log_level = "PROGRESS"
                    message   = 100
                }
            }
        }
        Remove-PSDrive -Name "Z"
    }
    catch {
        [pscustomobject]@{
            log_level = "ERROR"
            message   = $_.Exception.Message
        }
        
    }
}

$ConfirmCopyISO = {
    [bool](Get-ChildItem $(Join-Path -Path $using:Workspace -ChildPath $using:ISOFullName) -ErrorAction SilentlyContinue)
}

if ([bool](Test-WSMan -ComputerName $Computer -ErrorAction SilentlyContinue)) {
    try {
        Log-Info -Message "Connectivity with '$Computer' is verified."
        Log-Info -Message "Starting ISO copy from shared location"

        Invoke-Command -ComputerName $Computer -Credential $Credential -ScriptBlock $CopyISOWithProgress | Export-Csv -Path $Global:LogFileLocation -Encoding ASCII -Append -NoTypeInformation -Force
        $InvokeConfirmCopyISO = Invoke-Command -ComputerName $Computer -Credential $Credential -ScriptBlock $ConfirmCopyISO
        if ($InvokeConfirmCopyISO -eq $false) {
            Log-Error -Message "Failed to copy ISO file to the host '$Computer'."
            exit
        }
        Log-Info -Message "ISO copied successfully."
    }
    catch {
        #$_.Exception | Select-Object ErrorRecord, CommandName, Message | Format-List * -Force
        Log-Error -Message $("ErrorRecord : {0}, CommandName: {1}, Message: {2}" -f $_.Exception.ErrorRecord, $_.Exception.CommandName, $_.Exception.Message)
    }
}
else {
    Log-Error -Message "Can not connect to $Computer. Make sure the hostname is valid and PS-Remoting is allowed on the target machine."
}
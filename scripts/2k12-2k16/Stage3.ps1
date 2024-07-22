param(
    #Vsphare Params
    $VCenterAddress = "REPLACE_ME",
    $VCenterAuthUser = "REPLACE_ME",
    $VCenterAuthPass = "REPLACE_ME",
    $VMName = "REPLACE_ME",

    #GuestParams
    $Computer = "REPLACE_ME",
    $AuthUser = "REPLACE_ME",
    $AuthPassword = "REPLACE_ME",
    $ServiceBackupLocation = "REPLACE_ME", #eg: "\\192.168.31.90\share"
    $ISOFullName = "REPLACE_ME", #E.g. "SRV2016.STD.ENU.MAY2022.iso"
    $ImageIndex = 2,
    $Workspace = "C:\Temp",
    $ToEmail = "",
    $FromEmail = ""
)

$ErrorActionPreference = "Stop"

$HostReachableRetry = 5
$AllowedUpgradeVersion = "2016"
$SecurePassword = $AuthPassword | ConvertTo-SecureString -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential($AuthUser, $SecurePassword)

$VCenterAuthPass = $VCenterAuthPass | ConvertTo-SecureString -AsPlainText -Force
$VcenterCredentials = New-Object System.Management.Automation.PSCredential($VCenterAuthUser, $VCenterAuthPass)

<#
$VMOldName = "{0}_OLD" -f $VMName
$VMCloneName = "{0}_CLONE" -f $VMName
$DeleteDate = ((get-date).AddDays(14)).ToString("ddMMMM")
$VMDeleteName = "{0}_DeleteAfter_{1}" -f $VMName, $DeleteDate
#>

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

Function Log-Progress {
    param(
        [Parameter(ValueFromPipeline = $true)]
        [string]$Percentage
    ) 
    $logEntry = [pscustomobject]@{
        log_level = "PROGRESS"
        message   = $Percentage
    }
    $logEntry | Export-Csv -Path $Global:LogFileLocation -Encoding ASCII -Append -NoTypeInformation -Force
}

Start-Log -LogFileName "stage3.csv"

$PrepareInstallation = {
    try {
        $MountPath = Mount-DiskImage -ImagePath $(Join-Path -Path $using:Workspace -ChildPath $using:ISOFullName) -PassThru
        $DriveLetter = ($MountPath | Get-Volume).DriveLetter
        $DriveLetter = "$DriveLetter`:"
        $SetupLocation = "$DriveLetter\setup.exe"
        #Create upgrade script
        $CommandFile = "$($using:Workspace)\upgrade.cmd"
        Set-Content -Path $CommandFile -Value "@echo off" -Force
        Add-Content -Path $CommandFile -Value "start /wait $SetupLocation /auto upgrade /quiet /noreboot /dynamicupdate disable /compat ignorewarning /imageindex $using:ImageIndex" -Force
        Add-Content -Path $CommandFile -Value "echo %ERRORLEVEL% >$($using:Workspace)\log_upgrade.txt" -Force
    }
    catch {
        $_.Exception.Message
    }
}


$ConfirmInstallationPrep = {
    [bool](Get-ChildItem $(Join-Path -Path $using:Workspace -ChildPath "upgrade.cmd") -ErrorAction SilentlyContinue)
}


$RunUpgrade = {
    $CommandFile = "$($using:Workspace)\upgrade.cmd"
    cmd.exe /c "$CommandFile"
}

$ConfirmUpgrade = {
    $LogFile = "$($using:Workspace)\log_upgrade.txt"
    [int](Get-Content $LogFile) -eq 0
}

$ExportServices = {
    param(
        [string]$Filename
    )
    try {
        $ExportPath = Join-Path -Path "C:" -ChildPath "upgrade_temp" 
        [System.IO.Directory]::CreateDirectory($ExportPath)
        Get-Service | Select-Object Name, DisplayName, Status | Export-Csv $(Join-Path -Path $ExportPath -ChildPath $Filename) -NoTypeInformation 
    }
    catch {
        $_.Exception.Message
    }
}

$ExportSoftwares = {
    param(
        [string]$Filename
    )
    try {
        $ExportPath = Join-Path -Path "C:" -ChildPath "upgrade_temp" 
        [System.IO.Directory]::CreateDirectory($ExportPath)
        Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName } | Select-Object DisplayName | Export-Csv $(Join-Path -Path $ExportPath -ChildPath $Filename) -NoTypeInformation
    }
    catch {
        $_.Exception.Message
    }
}

$ExportDisks = {
    param(
        [string]$Filename
    )
    try {
        $ExportPath = Join-Path -Path "C:" -ChildPath "upgrade_temp" 
        [System.IO.Directory]::CreateDirectory($ExportPath)
        get-disk | select number, healthstatus, operationalstatus | Export-Csv $(Join-Path -Path $ExportPath -ChildPath $Filename) -NoTypeInformation
    }
    catch {
        $_.Exception.Message
    }
}

$CompareServiceDump = {
    param(
        [string]$FileNameBeforeUpgrade,
        [string]$FileNameAfterUpgrade
    )
    try {

        $ExportPath = Join-Path -Path "C:" -ChildPath "upgrade_temp" 
        $ServiceDumpBefore = Import-Csv  $(Join-Path -Path $ExportPath -ChildPath $FileNameBeforeUpgrade) -Delimiter ","
        $ServiceDumpAfter = Import-Csv  $(Join-Path -Path $ExportPath -ChildPath $FileNameAfterUpgrade) -Delimiter ","

        $ResultFile = Join-Path -Path  $ExportPath -ChildPath "ServiceAnalysis.txt"

        $NewServices = Compare-Object $ServiceDumpBefore $ServiceDumpAfter -Property Name | Where-Object { $_.SideIndicator -eq "=>" } | Select-Object -Expand Name
        $RemovedServices = Compare-Object $ServiceDumpBefore $ServiceDumpAfter -Property Name | Where-Object { $_.SideIndicator -eq "<=" } | Select-Object -Expand Name

        Set-Content -Path $ResultFile -Value "##############################################" -Force
        Add-Content -Path $ResultFile -Value "#            NEW SERVICES ADDED              #" -Force
        Add-Content -Path $ResultFile -Value "##############################################" -Force
        Add-Content -Path $ResultFile -Value $NewServices -Force

        Add-Content -Path $ResultFile -Value "##############################################" -Force
        Add-Content -Path $ResultFile -Value "#              SERVICES REMOVED              #" -Force
        Add-Content -Path $ResultFile -Value "##############################################" -Force
        Add-Content -Path $ResultFile -Value $RemovedServices -Force

        Add-Content -Path $ResultFile -Value "##############################################" -Force
        Add-Content -Path $ResultFile -Value "#            SERVICES WITH DIFERENCE         #" -Force
        Add-Content -Path $ResultFile -Value "##############################################" -Force

        foreach ($Item in $ServiceDumpBefore) {
            $Diff = $ServiceDumpAfter | Where-Object { $_.Name -eq $item.name -and $_.Status -ne $item.Status }
            if ($Diff) {
                $Data = $item | Select-Object Name, DisplayName, @{n = 'Status_Before'; e = { $Item.Status } }, @{n = 'Status_After'; e = { $Diff.Status } }
                $Data = "Name         :`t{0}`nDisplayName  :`t{1}`nStatus Before:`t{2}`nStatus After :`t{3}`n" -f $Data.Name, $Data.DisplayName, $Data.Status_Before, $Data.Status_After
                Add-Content -Path $ResultFile -Value $Data -Force              
            }
        }
    }
    catch {
        $_.Exception.Message
    }
}

$CompareSoftwareDump = {
    param(
        [string]$FileNameBeforeUpgrade,
        [string]$FileNameAfterUpgrade
    )
    try {

        $ExportPath = Join-Path -Path "C:" -ChildPath "upgrade_temp" 
        $SoftwaresDumpBefore = Import-Csv  $(Join-Path -Path $ExportPath -ChildPath $FileNameBeforeUpgrade) -Delimiter ","
        $SoftwaresDumpAfter = Import-Csv  $(Join-Path -Path $ExportPath -ChildPath $FileNameAfterUpgrade) -Delimiter ","

        $ResultFile = Join-Path -Path  $ExportPath -ChildPath "SoftwareAnalysis.txt"
        $NewSoftwares = $null
        $RemovedSoftwares = $null
        if ($SoftwaresDumpBefore -and $SoftwaresDumpAfter) {
            $NewSoftwares = Compare-Object $SoftwaresDumpBefore $SoftwaresDumpAfter -Property DisplayName | Where-Object { $_.SideIndicator -eq "=>" } | Select-Object -Expand DisplayName
            $RemovedSoftwares = Compare-Object $SoftwaresDumpBefore $SoftwaresDumpAfter -Property DisplayName | Where-Object { $_.SideIndicator -eq "<=" } | Select-Object -Expand DisplayName    
        }
        Set-Content -Path $ResultFile -Value "##############################################" -Force
        Add-Content -Path $ResultFile -Value "#            NEW SOFTWARES ADDED              #" -Force
        Add-Content -Path $ResultFile -Value "##############################################" -Force
        Add-Content -Path $ResultFile -Value $NewSoftwares -Force

        Add-Content -Path $ResultFile -Value "##############################################" -Force
        Add-Content -Path $ResultFile -Value "#              SOFTWARES REMOVED              #" -Force
        Add-Content -Path $ResultFile -Value "##############################################" -Force
        Add-Content -Path $ResultFile -Value $RemovedSoftwares -Force
    }
    catch {
        $_.Exception.Message
    }
}

$CompareDiskDump = {
    param(
        [string]$FileNameBeforeUpgrade,
        [string]$FileNameAfterUpgrade
    )
    try {

        $ExportPath = Join-Path -Path "C:" -ChildPath "upgrade_temp" 
        $DiskDumpBefore = Import-Csv  $(Join-Path -Path $ExportPath -ChildPath $FileNameBeforeUpgrade) -Delimiter ","
        $DiskDumpAfter = Import-Csv  $(Join-Path -Path $ExportPath -ChildPath $FileNameAfterUpgrade) -Delimiter ","

        $ResultFile = Join-Path -Path  $ExportPath -ChildPath "DiskAnalysis.txt"
        $NewDisks = $null
        $RemovedDisks = $null
        $DisksWithDifference = @()

        if ($DiskDumpBefore -and $DiskDumpAfter) {
            $NewDisks = Compare-Object $DiskDumpBefore $DiskDumpAfter -Property number | Where-Object { $_.SideIndicator -eq "=>" } | Select-Object -Expand number
            $RemovedDisks = Compare-Object $DiskDumpBefore $DiskDumpAfter -Property number | Where-Object { $_.SideIndicator -eq "<=" } | Select-Object -Expand number
            
            foreach ($Item in $DiskDumpBefore) {
                $Diff = $DiskDumpAfter | Where-Object { $_.number -eq $item.number -and $_.OperationalStatus -ne $item.OperationalStatus }
                if ($Diff) {
                    $Data = $item | Select-Object number, @{n = 'Status_Before'; e = { $Item.OperationalStatus } }, @{n = 'Status_After'; e = { $Diff.OperationalStatus } }
                    $DisksWithDifference += $Data
                }
            }
        }
       
        Set-Content -Path $ResultFile -Value "##############################################" -Force
        Add-Content -Path $ResultFile -Value "#            NEW DISKS ADDED                 #" -Force
        Add-Content -Path $ResultFile -Value "##############################################" -Force
        Add-Content -Path $ResultFile -Value $NewDisks -Force

        Add-Content -Path $ResultFile -Value "##############################################" -Force
        Add-Content -Path $ResultFile -Value "#              DISKS REMOVED                 #" -Force
        Add-Content -Path $ResultFile -Value "##############################################" -Force
        Add-Content -Path $ResultFile -Value $RemovedDisks -Force

        Add-Content -Path $ResultFile -Value "##############################################" -Force
        Add-Content -Path $ResultFile -Value "#            DISKS WITH DIFERENCE            #" -Force
        Add-Content -Path $ResultFile -Value "##############################################" -Force
        if ($DisksWithDifference) {
            foreach ($Data in $DisksWithDifference) {
                $Data = "Number  :`t{0}`nStatus Before:`t{1}`nStatus After :`t{2}`n" -f $Data.number, $Data.Status_Before, $Data.Status_After
                Add-Content -Path $ResultFile -Value $Data -Force              
            }
        }
    }
    catch {
        $_.Exception.Message
    }
}

function Send-StageStatusEmail{
    param (
        $FromEmail,
        $ToEmail,
        $Subject
    )

    $logs = Import-Csv $Global:LogFileLocation
    $MailMessage = '<!DOCTYPE html><html><head></head><body>'
    $MailMessage += '<p>Hello Team,</br></br>Your initiated stage-3 for Windows inplace upgrade is Completed. Please check the logs below and take appropriate action.</br></br></p>'
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

try {
    
    Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP $false -Confirm:$false
    Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false

    Log-Info -Message "Trying to connect to VIServer."
    Connect-VIServer -Server $VCenterAddress -Credential $VcenterCredentials
    Log-Info -Message "Connected to VIServer."

    [array]$Computers = $Computer -split "," | Select-Object -Unique
    if ($Computers.Length -gt 5) {
        Log-Error -Message "Only 5 servers are allowed for upgrade."
        exit
    }

    $ServersReadyForUpgrade = @()
    foreach ($Computer in $Computers) {
        try {
            $Computer = $Computer.trim()
            $VMName = $Computer
            $VMOldName = "{0}_OLD" -f $VMName
            $VMCloneName = "{0}_CLONE" -f $VMName
            
            Log-info -Message "Processing host '$Computer'"
            $VM = Get-VM -Name $VMName
            $VMProvisionedSpace = $VM | Select-Object -Expand ProvisionedSpaceGB
            $VMProvisionedSpace = [Math]::Round($VMProvisionedSpace, 2) #Round decimal to 2 digit
            #Log-Info -Message "Allocated disk space of VM is $VMProvisionedSpace GB."
            $VMClusterObject = Get-VM -Name $VMName | Select-Object @{n = 'Cluster'; e = { Get-Cluster -VM $_ } }
            $VMClusterName = $VMClusterObject.Cluster.Name
    
            $AvailableDataStore = $null

            $VMDataStores = $VM | Get-Datastore
            foreach ($vmds in $VMDataStores) {
                if ([Math]::Round($vmds.FreeSpaceGB, 2) -gt (2 * $VMProvisionedSpace )) {
                    $AvailableDataStore = $vmds.Name
                    break
                }
            }

            if ($null -eq $AvailableDataStore) {
        
                $ClusterDataStores = Get-Cluster -Name $VMClusterName | Get-Datastore | Where-Object { $_.capacitygb -gt 300 } -ErrorAction SilentlyContinue

                #$ClusterDataStores = get-cluster -vm $VM | Get-Datastore | Where-Object { $_.capacitygb -gt 300 } -ErrorAction SilentlyContinue 
                if (-not $ClusterDataStores) {
                    Log-Error "There is no datastore assigned to the cluster."
                    continue
                }

                Log-Info -Message "Checking if any Datastore in cluster has enough space available."
    
                foreach ($DataStore in $ClusterDataStores) {
                    if ([Math]::Round($DataStore.FreeSpaceGB, 2) -gt (2 * $VMProvisionedSpace )) {
                        $AvailableDataStore = $DataStore.Name
                        break
                    }
                }
            }

            if (-not $AvailableDataStore) {
                Log-Error -Message "No Datastore in cluster has required capacity to create a clone of the vm."
                continue
            }
            Log-Info -Message "Datastore '$AvailableDataStore' has required capacity to create a clone."

            Log-Info -Message "Turing VM off."
            Stop-VM -VM $VMName -Confirm:$false -ErrorAction SilentlyContinue #erroraction is used to avoid error if machine is already in turned off state.
    
            #Rename vm
            Log-Info -Message "Renaming VM '$VMName' to '$VMOldName'."
            Get-VM -Name $VMName | Set-VM -Name $VMOldName -Confirm:$false | Out-Null

            #Create a clone

            $DataStoreObject = Get-Datastore -Name $AvailableDataStore
            if ($DataStoreObject.Length -gt 1) {
                $DataStoreObject = $DataStoreObject[0]
            } 
            Log-Info -Message "Creating a clone from '$VMOldName' with name '$VMCloneName'."
            New-VM -VM $VMOldName -Name $VMCloneName -Datastore $DataStoreObject -ResourcePool $VMClusterName -Confirm:$false | Out-Null
    
            #Starting Cloned VM
            Log-Info -Message "Powering on VM '$VMCloneName'."
            Start-VM -VM $VMCloneName -Confirm:$false | Out-Null

            Log-Info -Message "Waiting for VM to be reachable."

            #It might take some time for server to get ready to take new commands.
            while ($HostReachableRetry -ge 0) {
                if ([bool](Test-WSMan -ComputerName $VMName -ErrorAction SilentlyContinue)) {
                    #Assuming cloned machine is discoverable on network with same hostname
                    break
                }
                Start-Sleep -Seconds 15
                $HostReachableRetry = $HostReachableRetry - 1  
            }

            if ([bool](Test-WSMan -ComputerName $Computer -ErrorAction SilentlyContinue)) {
                Log-Info -Message "Connectivity with '$Computer' is verified."
                Log-Info -Message "Preparing host for the upgrade."

                $InvokePrepereInstallation = Invoke-Command -ComputerName $Computer -Credential $Credential -ScriptBlock $PrepareInstallation
                $InvokeConfirmPrepereInstallation = Invoke-Command -ComputerName $Computer -Credential $Credential -ScriptBlock $ConfirmInstallationPrep
                if ($InvokeConfirmPrepereInstallation -eq $false) {
                    Log-Error -Message "Failed to prepare host '$Computer' for upgrade."
                    Log-Error -Message $InvokePrepereInstallation
                    continue
                }
    
                $ServersReadyForUpgrade += $Computer
            }
            else {
                Log-Error -Message "Can not connect to $Computer. Make sure the hostname is valid and PS-Remoting is allowed on the target machine."
            }
        }
        catch {
            Log-Error -Message $("ErrorRecord : {0}, CommandName: {1}, Message: {2}" -f $_.Exception.ErrorRecord, $_.Exception.CommandName, $_.Exception.Message)        
        }
    }

    #invoke server upgrade command on multiple servers
    Log-Info -Message $("Servers ready for upgrade - '{0}'. Running upgrade process.." -f $($ServersReadyForUpgrade -join ","))
    Invoke-Command -ComputerName $ServersReadyForUpgrade -Credential $Credential -ScriptBlock $RunUpgrade
    $InvokeConfirmUpgrade = Invoke-Command -ComputerName $ServersReadyForUpgrade -Credential $Credential -ScriptBlock $ConfirmUpgrade
    
    $SrvUpgradedSuccessfully = $InvokeConfirmUpgrade | Where-Object { $_ -eq $true } | Select-Object -ExpandProperty PSComputerName
    $SrvFailedToUpgrade = $InvokeConfirmUpgrade | Where-Object { $_ -eq $false } | Select-Object -ExpandProperty PSComputerName
    
    if ($SrvFailedToUpgrade) {
        Log-Error -Message $("Servers upgrade failed for host(s) - '{0}'. Check logs on server for troubleshooting." -f $($SrvFailedToUpgrade -join ","))
    }

    #Restart all successful servers at once
    Log-Info -Message $("Upgrade process completed for '{0}'. Restarting servers to apply upgrade." -f $($SrvUpgradedSuccessfully -join ","))
    Restart-Computer -ComputerName $SrvUpgradedSuccessfully -Credential $Credential -Protocol wsman -Force -Wait -For PowerShell

    #while waiting for servers to come back online perform retry and create analysis report.
    while ($HostReachableRetry -ge 0) {
        try {
            $InvokeExportService = Invoke-Command -ComputerName $SrvUpgradedSuccessfully -Credential $Credential -ScriptBlock $ExportServices -ArgumentList  "services_post_upgrade.csv"
            $InvokeExportSoftwares = Invoke-Command -ComputerName $SrvUpgradedSuccessfully -Credential $Credential -ScriptBlock $ExportSoftwares -ArgumentList  "softwares_post_upgrade.csv"
            $InvokeExportDisks = Invoke-Command -ComputerName $SrvUpgradedSuccessfully -Credential $Credential -ScriptBlock $ExportDisks -ArgumentList  "disks_post_upgrade.csv"
            break
        }
        catch {
            Start-Sleep -Seconds 60
            $HostReachableRetry = $HostReachableRetry - 1
        }
    }
    
    "Servers upgraded successfully: {0}" -f $($SrvUpgradedSuccessfully -join ",") | Out-File  -FilePath (Join-Path $PSScriptRoot "SRV_UPGRADE_STATUS.txt") -Force
    foreach ($Computer in $SrvUpgradedSuccessfully) {
        $VMName = $Computer
        $VMOldName = "{0}_OLD" -f $Computer
        $VMCloneName = "{0}_CLONE" -f $Computer
        $DeleteDate = ((get-date).AddDays(14)).ToString("ddMMMM")
        $VMDeleteName = "{0}_DeleteAfter_{1}" -f $Computer, $DeleteDate


        Log-Info -Message "Host '$Computer' has been upgraded."
        
        Log-Info -Message "Generating service dump analysis report."
        Invoke-Command -ComputerName $Computer -Credential $Credential -ScriptBlock $CompareServiceDump -ArgumentList ("services_pre_upgrade.csv", "services_post_upgrade.csv") -ErrorAction SilentlyContinue
        
        Log-Info -Message "Generating installed software dump analysis report."
        Invoke-Command -ComputerName $Computer -Credential $Credential -ScriptBlock $CompareSoftwareDump -ArgumentList ("softwares_pre_upgrade.csv", "softwares_post_upgrade.csv") -ErrorAction SilentlyContinue

        Log-Info -Message "Generating disk information dump analysis report."
        Invoke-Command -ComputerName $Computer -Credential $Credential -ScriptBlock $CompareDiskDump -ArgumentList ("disks_pre_upgrade.csv", "disks_post_upgrade.csv") -ErrorAction SilentlyContinue
        
        Log-Info -Message "Renaming VM '$VMCloneName' to '$VMName'."
        Get-VM -Name $VMCloneName | Set-VM -Name $VMName -Confirm:$false

        Log-Info -Message "Renaming VM '$VMOldName' to '$VMDeleteName'."
        Get-VM -Name $VMOldName | Set-VM -Name $VMDeleteName -Confirm:$false
    }
}
catch {
    Log-Error -Message $("ErrorRecord : {0}, CommandName: {1}, Message: {2}" -f $_.Exception.ErrorRecord, $_.Exception.CommandName, $_.Exception.Message)
}
finally{
    Log-Info "Sending email."
    Send-StageStatusEmail -FromEmail $FromEmail -ToEmail $($ToEmail -split ",") -Subject "Windows Upgrade Status: Stage 3"
}
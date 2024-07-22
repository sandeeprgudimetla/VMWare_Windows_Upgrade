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
    $LocalUserAccount = "REPLACE_ME",
    $LocalUserAccountPassword = "REPLACE_ME"
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

$VCenterAuthPass = $VCenterAuthPass | ConvertTo-SecureString -AsPlainText -Force
$VcenterCredentials = New-Object System.Management.Automation.PSCredential($VCenterAuthUser, $VCenterAuthPass)
#Start-Log -LogFileName "stage1.log"
Start-Log -LogFileName "stage1.csv"

$ExportServices = {
    param(
        [string]$Filename
    )
    try {
        $ExportPath = Join-Path -Path "C:" -ChildPath "upgrade_temp"
        [System.IO.Directory]::CreateDirectory($ExportPath)
        Get-Service | Select-Object Name, DisplayName, Status | Export-Csv $(Join-Path -Path $ExportPath -ChildPath $Filename) -NoTypeInformation 
        
        @{"Status" = "Success" }
    }
    catch {
        @{
            "Status" = "Failed"
            "Reason" = $($_.Exception.Message)
        }  
    }
}

$NewLocalUser = {
    try {
        if (-not [bool](Get-Command *LocalUser*)) {
        
            $ObjectOU = [ADSI]"WinNT://$env:COMPUTERNAME"
            $ObjUser = $ObjectOU.Create("User", $using:LocalUserAccount)
            $ObjUser.SetPassword($using:LocalUserAccountPassword)
            $ObjUser.Put("UserFlags", ($ObjUser.Properties["UserFlags"].Value -bor 0x10000)) # 0x10000 is the flag for "PasswordNeverExpires"
            $ObjUser.SetInfo()

            ##Add to local admin group
            $AdminGroup = [ADSI]"WinNT://$env:COMPUTERNAME/Administrators,group"
            $AdminGroup.Add("WinNT://$env:COMPUTERNAME/$($using:LocalUserAccount)")
        }
        else {
            $LocalUserAccountPassword = $using:LocalUserAccountPassword | ConvertTo-SecureString -AsPlainText -Force
            New-LocalUser -Name $using:LocalUserAccount -Password $LocalUserAccountPassword -PasswordNeverExpires | Out-Null
            Add-LocalGroupMember -Group "Administrators" -Member $using:LocalUserAccount
        }
    }
    catch {
        $_.Exception.Message
    }
}

$ConfirmNewLocalUser = {
    $localUsers = [ADSI]"WinNT://$env:COMPUTERNAME"
    [bool]($localUsers.Children | Where-Object { $_.SchemaClassName -eq 'user' -and $_.Name -eq $using:LocalUserAccount })
}

$GetCDriveFreeSpace = {
    [Math]::Round(((get-volume -DriveLetter "c" |Select-Object -expand SizeRemaining) / 1gb), 2)
}

function Check-NonVMXNetAdapter {
    param (
        $VMName
    )
    $NonVmxnetadapters = Get-NetworkAdapter -VM $VMName | Where-Object {$_.Type -ne 'VMXNet3'}
    [bool]$NonVmxnetadapters
}

function Check-RDMAssociation {
    param (
        $VMName
    )
    $Disks = Get-VM -Name $VMName | Get-HardDisk | Where-Object {$_.DiskType -eq 'RawPhysical'}
    [bool]$Disks 
}

try {
    #Assuming using pre-req is run and powercli is installed.
    Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP $false -Confirm:$false 
    Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false
    
    #Vsphare operations
    Log-Info -Message "Trying to connect to VIServer."
    Connect-VIServer -Server $VCenterAddress -Credential $VcenterCredentials
    Log-Info -Message "Connected to VIServer."

    Log-Info -Message "Getting VM's allocated disk space."
    $VM = Get-VM -Name $VMName
    $VMProvisionedSpace = $VM | Select-Object -Expand ProvisionedSpaceGB
    $VMProvisionedSpace = [Math]::Round($VMProvisionedSpace, 2) #Round decimal to 2 digit
    Log-Info -Message "Allocated disk space of VM is $VMProvisionedSpace GB."
    
    $AvailableDataStore = $null
    $VMDataStores = $VM | Get-Datastore
    foreach ($vmds in $VMDataStores) {
        if ([Math]::Round($vmds.FreeSpaceGB, 2) -gt (2 * $VMProvisionedSpace )) {
            $AvailableDataStore = $vmds.Name
            break
        }
    }

    if ($null -eq $AvailableDataStore) {
        #$VMClusterObject = Get-VM -Name $VMName | Select-Object @{n = 'Cluster'; e = { Get-Cluster -VM $_ } }
        $ClusterDataStores = get-cluster -vm $VM | Get-Datastore | Where-Object { $_.capacitygb -gt 300 } -ErrorAction SilentlyContinue 
        if (-not $ClusterDataStores) {
            Log-Error "There is no datastore assigned to the cluster."
            exit
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
        exit
    }

    Log-Info -Message "Datastore '$AvailableDataStore' has required capacity to create a clone."
    #Guest Operations
    if ([bool](Test-WSMan -ComputerName $Computer -ErrorAction SilentlyContinue)) {
        Log-Info -Message "Connectivity with '$Computer' is verified."
        
        Log-Info "Checking if user already present on the host."
        $InvokeConfirmNewLocalUser = Invoke-Command -ComputerName $Computer -Credential $Credential -ScriptBlock $ConfirmNewLocalUser
        if ($InvokeConfirmNewLocalUser -eq $false) {
            Log-Info -Message "User does not exist. Creating new local user account."
            $InvokeNewLocalUser = Invoke-Command -ComputerName $Computer -Credential $Credential -ScriptBlock $NewLocalUser
        }
        
        $InvokeConfirmNewLocalUser = Invoke-Command -ComputerName $Computer -Credential $Credential -ScriptBlock $ConfirmNewLocalUser
        if ($InvokeConfirmNewLocalUser -eq $false) {
            Log-Error -Message "Failed to create loca admin account."
            Log-Error -Message $InvokeNewLocalUser
            exit
        }


        Log-Info "Checking host for available free space in C drive."
        $InvokeGetCDriveFreeSpace = Invoke-Command -ComputerName $Computer -Credential $Credential -ScriptBlock $GetCDriveFreeSpace
        if ($InvokeGetCDriveFreeSpace -lt 40){
            Log-Error "Host '$Computer' does not have required space in C drive" 
            exit
        }
        Log-Info "Host '$Computer' has required space in C drive"

        Log-Info "Checking host for supported(VMXNet3) network adapter(s)."
        if(Check-NonVMXNetAdapter -VMName $Computer){
            Log-Error "Host '$Computer' has one or more unsupported network adapter(s)." 
            exit
        }
        Log-Info "Host '$Computer' has supported network adapter(s)."

        #Check-RDMAssociation
        Log-Info "Checking host '$Computer' for RDM Luns."
        if(Check-RDMAssociation -VMName $Computer){
            Log-Error "Host '$Computer' RDM Luns." 
            exit
        }
        Log-Info "Host '$Computer' does not have RDM Luns."

        Log-Info -Message "Taking Serevice dump."
        $InvokeExportService = Invoke-Command -ComputerName $Computer -Credential $Credential -ScriptBlock $ExportServices -ArgumentList "Pre_Upgrade.csv"
        if ($InvokeExportService.Status -eq "Failed") {
            Log-Error -Message "Failed to take service dump."
            Log-Error -Message $InvokeExportService.Reason
        }

        Log-Info -Message "Service dump has been taken."
    }
    else {
        Log-Error -Message "Can not connect to $Computer. Make sure the hostname is valid and PS-Remoting is allowed on the target machine."
        exit
    }
}
catch {
    #$_.Exception | Select-Object ErrorRecord, CommandName, Message | Format-List * -Force
    Log-Error -Message $("ErrorRecord : {0}, CommandName: {1}, Message: {2}" -f $_.Exception.ErrorRecord, $_.Exception.CommandName, $_.Exception.Message)
}
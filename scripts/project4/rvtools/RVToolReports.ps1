param(
    #Vsphare Params
    $VCenterAddress = "REPLACE_ME",
    $VCenterAuthUser = "REPLACE_ME",
    $VCenterAuthPass = "REPLACE_ME",
    $ToEmail = "gpk.pandey.neeraj@gmail.com",
    $FromEmail = "no_reply@notification.com"
)

$SMTP_USER = "abc@gmail.com"
$SMTP_PASS = "abc"

$REF_SHEETS = @(
    "vInfo",
    "vCPU",
    "vMemory",
    "vDisk",
    "vPartition",
    "vNetwork",
    "vCD",
    "vUSB",
    "vSnapshot",
    "vTools",
    "vSource",
    "vRP",
    "vCluster",
    "vHost",
    "vHBA",
    "vNIC",
    "vSwitch",
    "vPort",
    "dvSwitch",
    "dvPort",
    "vSC_VMK",
    "vDatastore",
    "vMultiPath",
    "vLicense",
    "vFileInfo",
    "vHealth",
    "vMetaData"
)

$REF_COLUMNS = @{
    vInfo      = @("VM", "Powerstate", "Template", "SRM Placeholder", "Config status", "DNS Name", "Connection state", "Guest state", "Heartbeat", "Consolidation Needed", "PowerOn", "Suspend time", "Creation date", "Change Version", "CPUs", "Memory", "NICs", "Disks", "Total disk capacity MiB", "min Required EVC Mode Key", "Latency Sensitivity", "EnableUUID", "CBT", "Primary IP Address", "Network #1", "Network #2", "Network #3", "Network #4", "Network #5", "Network #6", "Network #7", "Network #8", "Num Monitors", "Video Ram KiB", "Resource pool", "Folder ID", "Folder", "vApp", "DAS protection", "FT State", "FT Role", "FT Latency", "FT Bandwidth", "FT Sec. Latency", "Provisioned MiB", "In Use MiB", "Unshared MiB", "HA Restart Priority", "HA Isolation Response", "HA VM Monitoring", "Cluster rule(s)", "Cluster rule name(s)", "Boot Required", "Boot delay", "Boot retry delay", "Boot retry enabled", "Boot BIOS setup", "Reboot PowerOff", "EFI Secure boot", "Firmware", "HW version", "HW upgrade status", "HW upgrade policy", "HW target", "Path", "Log directory", "Snapshot directory", "Suspend directory", "Annotation", "app role", "app_role", "application", "application owner", "application_owner", "application_support", "compliance", "cost center", "costcenter", "data_classification", "dd_auto_discovery", "division", "environment", "finance_contact", "group", "group_beneficiary", "hfm_entity", "infrastructure_support", "LastBackupStatus-com.dellemc.avamar", "LastSuccessfulBackup-com.dellemc.avamar", "Name", "Owner", "project_number", "rightsizing_exception", "service account", "snow_support", "source_datacenter", "tranche", "VSphereExtensionUtil.SNAPSHOT_TAG", "Datacenter", "Cluster", "Host", "OS according to the configuration file", "OS according to the VMware Tools", "VM ID", "SMBIOS UUID", "VM UUID", "VI SDK Server type", "VI SDK API Version", "VI SDK Server", "VI SDK UUID")
    vCPU       = @("VM", "Powerstate", "Template", "SRM Placeholder", "CPUs", "Sockets", "Cores p/s", "Max", "Overall", "Level", "Shares", "Reservation", "Entitlement", "DRS Entitlement", "Limit", "Hot Add", "Hot Remove", "Numa Hotadd Exposed", "Annotation", "app role", "app_role", "application", "application owner", "application_owner", "application_support", "compliance", "cost center", "costcenter", "data_classification", "dd_auto_discovery", "division", "environment", "finance_contact", "group", "group_beneficiary", "hfm_entity", "infrastructure_support", "LastBackupStatus-com.dellemc.avamar", "LastSuccessfulBackup-com.dellemc.avamar", "Name", "Owner", "project_number", "rightsizing_exception", "service account", "snow_support", "source_datacenter", "tranche", "VSphereExtensionUtil.SNAPSHOT_TAG", "Datacenter", "Cluster", "Host", "Folder", "OS according to the configuration file", "OS according to the VMware Tools", "VM ID", "VM UUID", "VI SDK Server", "VI SDK UUID")
    vMemory    = @("VM", "Powerstate", "Template", "SRM Placeholder", "Size MiB", "Memory Reservation Locked To Max", "Overhead", "Max", "Consumed", "Consumed Overhead", "Private", "Shared", "Swapped", "Ballooned", "Active", "Entitlement", "DRS Entitlement", "Level", "Shares", "Reservation", "Limit", "Hot Add", "Annotation", "app role", "app_role", "application", "application owner", "application_owner", "application_support", "compliance", "cost center", "costcenter", "data_classification", "dd_auto_discovery", "division", "environment", "finance_contact", "group", "group_beneficiary", "hfm_entity", "infrastructure_support", "LastBackupStatus-com.dellemc.avamar", "LastSuccessfulBackup-com.dellemc.avamar", "Name", "Owner", "project_number", "rightsizing_exception", "service account", "snow_support", "source_datacenter", "tranche", "VSphereExtensionUtil.SNAPSHOT_TAG", "Datacenter", "Cluster", "Host", "Folder", "OS according to the configuration file", "OS according to the VMware Tools", "VM ID", "VM UUID", "VI SDK Server", "VI SDK UUID")
    vDisk      = @("VM", "Powerstate", "Template", "SRM Placeholder", "Disk", "Disk Key", "Disk UUID", "Disk Path", "Capacity MiB", "Raw", "Disk Mode", "Sharing mode", "Thin", "Eagerly Scrub", "Split", "Write Through", "Level", "Shares", "Reservation", "Limit", "Controller", "Label", "SCSI Unit #", "Unit #", "Shared Bus", "Path", "Raw LUN ID", "Raw Comp. Mode", "Internal Sort Column", "Annotation", "app role", "app_role", "application", "application owner", "application_owner", "application_support", "compliance", "cost center", "costcenter", "data_classification", "dd_auto_discovery", "division", "environment", "finance_contact", "group", "group_beneficiary", "hfm_entity", "infrastructure_support", "LastBackupStatus-com.dellemc.avamar", "LastSuccessfulBackup-com.dellemc.avamar", "Name", "Owner", "project_number", "rightsizing_exception", "service account", "snow_support", "source_datacenter", "tranche", "VSphereExtensionUtil.SNAPSHOT_TAG", "Datacenter", "Cluster", "Host", "Folder", "OS according to the configuration file", "OS according to the VMware Tools", "VM ID", "VM UUID", "VI SDK Server", "VI SDK UUID")
    vPartition = @("VM", "Powerstate", "Template", "SRM Placeholder", "Disk Key", "Disk", "Capacity MiB", "Consumed MiB", "Free MiB", "Free %", "Internal Sort Column", "Annotation", "app role", "app_role", "application", "application owner", "application_owner", "application_support", "compliance", "cost center", "costcenter", "data_classification", "dd_auto_discovery", "division", "environment", "finance_contact", "group", "group_beneficiary", "hfm_entity", "infrastructure_support", "LastBackupStatus-com.dellemc.avamar", "LastSuccessfulBackup-com.dellemc.avamar", "Name", "Owner", "project_number", "rightsizing_exception", "service account", "snow_support", "source_datacenter", "tranche", "VSphereExtensionUtil.SNAPSHOT_TAG", "Datacenter", "Cluster", "Host", "Folder", "OS according to the configuration file", "OS according to the VMware Tools", "VM ID", "VM UUID", "VI SDK Server", "VI SDK UUID")
    vNetwork   = @("VM", "Powerstate", "Template", "SRM Placeholder", "NIC label", "Adapter", "Network", "Switch", "Connected", "Starts Connected", "Mac Address", "Type", "IPv4 Address", "IPv6 Address", "Direct Path IO", "Internal Sort Column", "Annotation", "app role", "app_role", "application", "application owner", "application_owner", "application_support", "compliance", "cost center", "costcenter", "data_classification", "dd_auto_discovery", "division", "environment", "finance_contact", "group", "group_beneficiary", "hfm_entity", "infrastructure_support", "LastBackupStatus-com.dellemc.avamar", "LastSuccessfulBackup-com.dellemc.avamar", "Name", "Owner", "project_number", "rightsizing_exception", "service account", "snow_support", "source_datacenter", "tranche", "VSphereExtensionUtil.SNAPSHOT_TAG", "Datacenter", "Cluster", "Host", "Folder", "OS according to the configuration file", "OS according to the VMware Tools", "VM ID", "VM UUID", "VI SDK Server", "VI SDK UUID")
    vCD        = @("VM", "Powerstate", "Template", "SRM Placeholder", "Device Node", "Connected", "Starts Connected", "Device Type", "Annotation", "app role", "app_role", "application", "application owner", "application_owner", "application_support", "compliance", "cost center", "costcenter", "data_classification", "dd_auto_discovery", "division", "environment", "finance_contact", "group", "group_beneficiary", "hfm_entity", "infrastructure_support", "LastBackupStatus-com.dellemc.avamar", "LastSuccessfulBackup-com.dellemc.avamar", "Name", "Owner", "project_number", "rightsizing_exception", "service account", "snow_support", "source_datacenter", "tranche", "VSphereExtensionUtil.SNAPSHOT_TAG", "Datacenter", "Cluster", "Host", "Folder", "OS according to the configuration file", "OS according to the VMware Tools", "VMRef", "VM ID", "VM UUID", "VI SDK Server", "VI SDK UUID")
    vUSB       = @("VM", "Powerstate", "Template", "SRM Placeholder", "Device Node", "Device Type", "Connected", "Family", "Speed", "EHCI enabled", "Auto connect", "Bus number", "Unit number", "Annotation", "app role", "app_role", "application", "application owner", "application_owner", "application_support", "compliance", "cost center", "costcenter", "data_classification", "dd_auto_discovery", "division", "environment", "finance_contact", "group", "group_beneficiary", "hfm_entity", "infrastructure_support", "LastBackupStatus-com.dellemc.avamar", "LastSuccessfulBackup-com.dellemc.avamar", "Name", "Owner", "project_number", "rightsizing_exception", "service account", "snow_support", "source_datacenter", "tranche", "VSphereExtensionUtil.SNAPSHOT_TAG", "Datacenter", "Cluster", "Host", "Folder", "OS according to the configuration file", "OS according to the VMware tools", "VMRef", "VM ID", "VM UUID", "VI SDK Server", "VI SDK UUID")
    vSnapshot  = @("VM", "Powerstate", "Name", "Description", "Date / time", "Filename", "Size MiB (vmsn)", "Size MiB (total)", "Quiesced", "State", "Annotation", "app role", "app_role", "application", "application owner", "application_owner", "application_support", "compliance", "cost center", "costcenter", "data_classification", "dd_auto_discovery", "division", "environment", "finance_contact", "group", "group_beneficiary", "hfm_entity", "infrastructure_support", "LastBackupStatus-com.dellemc.avamar", "LastSuccessfulBackup-com.dellemc.avamar", "Name", "Owner", "project_number", "rightsizing_exception", "service account", "snow_support", "source_datacenter", "tranche", "VSphereExtensionUtil.SNAPSHOT_TAG", "Datacenter", "Cluster", "Host", "Folder", "OS according to the configuration file", "OS according to the VMware Tools", "VM ID", "VM UUID", "VI SDK Server", "VI SDK UUID")
    vTools     = @("VM", "Powerstate", "Template", "SRM Placeholder", "VM Version", "Tools", "Tools Version", "Required Version", "Upgradeable", "Upgrade Policy", "Sync time", "App status", "Heartbeat status", "Kernel Crash state", "Operation Ready", "State change support", "Interactive Guest", "Annotation", "app role", "app_role", "application", "application owner", "application_owner", "application_support", "compliance", "cost center", "costcenter", "data_classification", "dd_auto_discovery", "division", "environment", "finance_contact", "group", "group_beneficiary", "hfm_entity", "infrastructure_support", "LastBackupStatus-com.dellemc.avamar", "LastSuccessfulBackup-com.dellemc.avamar", "Name", "Owner", "project_number", "rightsizing_exception", "service account", "snow_support", "source_datacenter", "tranche", "VSphereExtensionUtil.SNAPSHOT_TAG", "Datacenter", "Cluster", "Host", "Folder", "OS according to the configuration file", "OS according to the VMware Tools", "VMRef", "VM ID", "VM UUID", "VI SDK Server", "VI SDK UUID")
    vSource    = @("Name", "OS type", "API type", "API version", "Version", "Patch level", "Build", "Fullname", "Product name", "Product version", "Product line", "Vendor", "VI SDK Server", "VI SDK UUID")
    vRP        = @("Resource Pool name", "Resource Pool path", "Status", "# VMs total", "# VMs", "# vCPUs", "CPU limit", "CPU overheadLimit", "CPU reservation", "CPU level", "CPU shares", "CPU expandableReservation", "CPU maxUsage", "CPU overallUsage", "CPU reservationUsed", "CPU reservationUsedForVm", "CPU unreservedForPool", "CPU unreservedForVm", "Mem Configured", "Mem limit", "Mem overheadLimit", "Mem reservation", "Mem level", "Mem shares", "Mem expandableReservation", "Mem maxUsage", "Mem overallUsage", "Mem reservationUsed", "Mem reservationUsedForVm", "Mem unreservedForPool", "Mem unreservedForVm", "QS overallCpuDemand", "QS overallCpuUsage", "QS staticCpuEntitlement", "QS distributedCpuEntitlement", "QS balloonedMemory", "QS compressedMemory", "QS consumedOverheadMemory", "QS distributedMemoryEntitlement", "QS guestMemoryUsage", "QS hostMemoryUsage", "QS overheadMemory", "QS privateMemory", "QS sharedMemory", "QS staticMemoryEntitlement", "QS swappedMemory", "Object ID", "VI SDK Server", "VI SDK UUID")
    vCluster   = @("Name", "Config status", "OverallStatus", "NumHosts", "numEffectiveHosts", "TotalCpu", "NumCpuCores", "NumCpuThreads", "Effective Cpu", "TotalMemory", "Effective Memory", "Num VMotions", "HA enabled", "Failover Level", "AdmissionControlEnabled", "Host monitoring", "HB Datastore Candidate Policy", "Isolation Response", "Restart Priority", "Cluster Settings", "Max Failures", "Max Failure Window", "Failure Interval", "Min Up Time", "VM Monitoring", "DRS enabled", "DRS default VM behavior", "DRS vmotion rate", "DPM enabled", "DPM default behavior", "DPM Host Power Action Rate", "Object ID", "com.vmware.vcenter.cluster.edrs.upgradeHostAdded", "VI SDK Server", "VI SDK UUID")
    vHost      = @("Host", "Datacenter", "Cluster", "Config status", "in Maintenance Mode", "in Quarantine Mode", "vSAN Fault Domain Name", "CPU Model", "Speed", "HT Available", "HT Active", "# CPU", "Cores per CPU", "# Cores", "CPU usage %", "# Memory", "Memory usage %", "Console", "# NICs", "# HBAs", "# VMs total", "# VMs", "VMs per Core", "# vCPUs", "vCPUs per Core", "vRAM", "VM Used memory", "VM Memory Swapped", "VM Memory Ballooned", "VMotion support", "Storage VMotion support", "Current EVC", "Max EVC", "Assigned License(s)", "ATS Heartbeat", "ATS Locking", "Current CPU power man. policy", "Supported CPU power man.", "Host Power Policy", "ESX Version", "Boot time", "DNS Servers", "DHCP", "Domain", "DNS Search Order", "NTP Server(s)", "NTPD running", "Time Zone", "Time Zone Name", "GMT Offset", "Vendor", "Model", "Serial number", "Service tag", "OEM specific string", "BIOS Vendor", "BIOS Version", "BIOS Date", "Object ID", "AutoDeploy.MachineIdentity", "LastBackupStatus-com.dellemc.avamar", "LastSuccessfulBackup-com.dellemc.avamar", "Remove this VM - Cannot Delete", "VSphereExtensionUtil.SNAPSHOT_TAG", "UUID", "VI SDK Server", "VI SDK UUID")
    vHBA       = @("Host", "Datacenter", "Cluster", "Device", "Type", "Status", "Bus", "Pci", "Driver", "Model", "WWN", "VI SDK Server", "VI SDK UUID")
    vNIC       = @("Host", "Datacenter", "Cluster", "Network Device", "Driver", "Speed", "Duplex", "MAC", "Switch", "Uplink port", "PCI", "WakeOn", "VI SDK Server", "VI SDK UUID")
    vSwitch    = @("Host", "Datacenter", "Cluster", "Switch", "# Ports", "Free Ports", "Promiscuous Mode", "Mac Changes", "Forged Transmits", "Traffic Shaping", "Width", "Peak", "Burst", "Policy", "Reverse Policy", "Notify Switch", "Rolling Order", "Offload", "TSO", "Zero Copy Xmit", "MTU", "VI SDK Server", "VI SDK UUID")
    vPort      = @("Host", "Datacenter", "Cluster", "Port Group", "Switch", "VLAN", "Promiscuous Mode", "Mac Changes", "Forged Transmits", "Traffic Shaping", "Width", "Peak", "Burst", "Policy", "Reverse Policy", "Notify Switch", "Rolling Order", "Offload", "TSO", "Zero Copy Xmit", "VI SDK Server", "VI SDK UUID")
    dvSwitch   = @("Switch", "Datacenter", "Name", "Vendor", "Version", "Description", "Created", "Host members", "Max Ports", "# Ports", "# VMs", "In Traffic Shaping", "In Avg", "In Peak", "In Burst", "Out Traffic Shaping", "Out Avg", "Out Peak", "Out Burst", "CDP Type", "CDP Operation", "LACP Name", "LACP Mode", "LACP Load Balance Alg.", "Max MTU", "Contact", "Admin Name", "Object ID", "VI SDK Server", "VI SDK UUID")
    dvPort     = @("Port", "Switch", "Type", "# Ports", "VLAN", "Speed", "Full Duplex", "Blocked", "Allow Promiscuous", "Mac Changes", "Active Uplink", "Standby Uplink", "Policy", "Forged Transmits", "In Traffic Shaping", "In Avg", "In Peak", "In Burst", "Out Traffic Shaping", "Out Avg", "Out Peak", "Out Burst", "Reverse Policy", "Notify Switch", "Rolling Order", "Check Beacon", "Live Port Moving", "Check Duplex", "Check Error %", "Check Speed", "Percentage", "Block Override", "Config Reset", "Shaping Override", "Vendor Config Override", "Sec. Policy Override", "Teaming Override", "Vlan Override", "Object ID", "VI SDK Server", "VI SDK UUID")
    vSC_VMK    = @("Host", "Datacenter", "Cluster", "Port Group", "Device", "Mac Address", "DHCP", "IP Address", "IP 6 Address", "Subnet mask", "Gateway", "IP 6 Gateway", "MTU", "VI SDK Server", "VI SDK UUID")
    vDatastore = @("Name", "Config status", "Address", "Accessible", "Type", "# VMs total", "# VMs", "Capacity MiB", "Provisioned MiB", "In Use MiB", "Free MiB", "Free %", "SIOC enabled", "SIOC Threshold", "# Hosts", "Hosts", "Cluster name", "Cluster capacity MiB", "Cluster free space MiB", "Block size", "Max Blocks", "# Extents", "Major Version", "Version", "VMFS Upgradeable", "MHA", "URL", "Object ID", "VI SDK Server", "VI SDK UUID")
    vMultiPath = @("Host", "Cluster", "Datacenter", "Datastore", "Disk", "Display name", "Policy", "Oper. State", "Path 1", "Path 1 state", "Path 2", "Path 2 state", "Path 3", "Path 3 state", "Path 4", "Path 4 state", "Path 5", "Path 5 state", "Path 6", "Path 6 state", "Path 7", "Path 7 state", "Path 8", "Path 8 state", "vStorage", "Queue depth", "Vendor", "Model", "Revision", "Level", "Serial #", "UUID", "Object ID", "VI SDK Server", "VI SDK UUID")
    vLicense   = @("Name", "Key", "Labels", "Cost Unit", "Total", "Used", "Expiration Date", "Features", "VI SDK Server", "VI SDK UUID")
    vFileInfo  = @("Friendly Path Name", "File Name", "File Type", "File Size in bytes", "Path", "Internal Sort Column", "VI SDK Server", "VI SDK UUID")
    vHealth    = @("Name", "Message", "Message type", "VI SDK Server", "VI SDK UUID")
    vMetaData  = @("RVTools major version", "RVTools version", "xlsx creation datetime", "Server")
}

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
        $MailMessage += '<p>Hello Team,</br></br>Your initiated RVTool report generation has been complted without final report. Please contact your administrator.</br></br></p>'
    }
    else {
        $MailMessage += '<p>Hello Team,</br></br>Your initiated RVTool report generation has been complted. Please check the logs below and take appropriate action.</br></br></p>'
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


$DateFormat = $(get-date -Format 'dd_MM_yyyy_hh_mm')
$RVToolReportFolderName = "RVToolReports_{0}" -f $DateFormat


$RVToolReportsPath = (Join-Path $PSScriptRoot $RVToolReportFolderName)
[System.IO.Directory]::CreateDirectory($RVToolReportsPath) | Out-Null
$RVToolCombinedReportFile = (Join-Path $RVToolReportsPath "CombinedReport_$DateFormat.xlsx")
$LogFileName = "rvtools_{0}.csv" -f $DateFormat

Try {
    Start-Log -LogFileName $LogFileName
    [array]$VCenterServers = $VCenterAddress -split ","
    Log-Info -Message "Checking if RVtools is available or not."
    if(Test-Path "C:\Program Files (x86)\Robware\RVTools\rvtools.exe"){
        Log-Info -Message "RVtools found on the server." 
        foreach ($server in $VCenterServers) {
            try {
                Log-Info -Message "Downloading report from '$server'." 
                Set-Location "C:\Program Files (x86)\Robware\RVTools"
                $ReportName = $server + "_" + $(get-date -Format MMddyyyy) + "_" + ".xlsx"
                .\rvtools.exe -u $VCenterAuthUser -p $VCenterAuthPass  -s $server -c ExportAll2xls -d $RVToolReportsPath -f $ReportName
                Start-Sleep -Seconds 5
            }
            catch {
                Log-Error -Message "Failed to download report from '$server'"
                Log-Error -Message $("ErrorRecord : {0}, CommandName: {1}, Message: {2}" -f $_.Exception.ErrorRecord, $_.Exception.CommandName, $_.Exception.Message)
            }
        }
        
        $FilesToProcess = Get-ChildItem $RVToolReportsPath -Filter '*.xlsx'
        if (-not $FilesToProcess) {
            Log-Error -Message "No report was found for processing."
        }
        else {
            foreach ($File in $FilesToProcess) {
                foreach ($worksheet in $REF_SHEETS) {
                    Log-Info -Message "Processing worksheet '$worksheet' from file '$File'."
                    try {
                        [array]$sheetdata = Import-Excel $File.FullName -WorksheetName $worksheet -HeaderName $($REF_COLUMNS[$worksheet] | select -Unique)
                        $sheetdata = $sheetdata[1..$sheetdata.Length] #remove first entry of the array as it will contain headers
                        $sheetdata | Export-Excel $RVToolCombinedReportFile -Append -WorksheetName $worksheet
                    }
                    catch [System.Management.Automation.RuntimeException] {
                        Log-Error -Message "Error while reading data from worksheet. Worksheet may not exist in the file."
                    }
                    catch {
                        Log-Error -Message "Error while processing worksheet."
                        Log-Error -Message $("ErrorRecord : {0}, CommandName: {1}, Message: {2}" -f $_.Exception.ErrorRecord, $_.Exception.CommandName, $_.Exception.Message)    
                    }
                }
            }
        }    
    }
    else{
        Log-Error -Message "RVTools not found on the server."
    }
}
catch {
    Log-Error -Message $("ErrorRecord : {0}, CommandName: {1}, Message: {2}" -f $_.Exception.ErrorRecord, $_.Exception.CommandName, $_.Exception.Message)
}
finally {
    Log-Info "Sending email."
    Send-StageStatusEmail -FromEmail $FromEmail -ToEmail $($ToEmail -split ",") -Subject "RVTool Report." -ReportPath $RVToolCombinedReportFile
}





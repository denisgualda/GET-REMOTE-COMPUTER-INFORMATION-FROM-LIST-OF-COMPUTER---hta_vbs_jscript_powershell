$OutputLog = "C:\PRE_PROVES_SCRIPTS\INFO_MONITOR_POWERSHELL\Machine_MainLog.txt" # Main log
$NotRespondingLog = "C:\PRE_PROVES_SCRIPTS\INFO_MONITOR_POWERSHELL\Machine_NoResponse_Log.txt" # Logging "unqueried" hosts

$ErrorActionPreference = "Stop" # Or add '-EA Stop' to Get-WmiObject queries
Clear-Host
#$Computer = "LT2B151067"
$Computer = $args[0]

        $computerSystem = get-wmiobject Win32_ComputerSystem -Computer $Computer
        $computerBIOS = get-wmiobject Win32_BIOS -Computer $Computer
        $LoggedOnUser = $ComputerSystem.UserName
        $Version = Get-WmiObject -Namespace "Root\CIMv2" `
            -Query "Select * from Win32_ComputerSystemProduct" `
            -computer $computer | select -ExpandProperty version
        $MonitorInfo = gwmi WmiMonitorID -Namespace root\wmi `
            -computername $Computer `
            | Select PSComputerName, `
                @{n="Model";e={[System.Text.Encoding]::ASCII.GetString(`
                    $_.UserFriendlyName -ne 00)}},
                @{n="Serial Number";e={[System.Text.Encoding]::ASCII.GetString(`
                    $_.SerialNumberID -ne 00)}}     


    #$Header = "System Information for: {0}" -f $computerSystem.Name

    # Outputting and logging header.
    write-host $Header -BackgroundColor DarkCyan
    $Header | Out-File -FilePath $OutputLog -Append -Encoding UTF8

    $Output = (@"
MODEL  : {4}
S/N    : {5}
"@) -f $computerSystem.Model,$LoggedOnUser, $computerBIOS.SerialNumber, $Version, `
       $MonitorInfo.Model, $MonitorInfo."Serial Number"

    # Ouputting and logging WMI data
    Write-Host $Output
    $Output | Out-File -FilePath $OutputLog -Append -Encoding UTF8

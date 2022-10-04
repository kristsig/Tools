### Collect useful info when a Windows Operating System Deployment has failed
Clear-Host

############ Editable Variables #############
$Filename_ThisScript = "Get-OSDFailInfo.ps1"
#Specific files that need to be found
$List_FileSearch = @(
    "smsts.log",
    "netsetup.log"#,
    #'*.ps1'
    #'*.log'
)

# Change thisto whatever IPs you would like to test if can be reached.
$List_Ping = @(
    "127.0.0.1",
    "1.1.1.1",
    "8.8.8.8"
)

#############################################

#Get current hosts serialnumber
$HW_Serialnumber = Get-WmiObject win32_bios | select Serialnumber
$HW_Serialnumber = $HW_Serialnumber.Serialnumber
Write-Host "HW_Serialnumber"
Write-Host $HW_Serialnumber

# UUID Local
Function Get-UUIDLocal{
    $UUID = ""
    try{
        $UUID = (Get-WmiObject -Class Win32_ComputerSystemProduct).UUID
    }Catch{
        try{
            $UUID = (Get-CimInstance -Class Win32_ComputerSystemProduct).UUID
        }catch{}
    }
    Return $UUID
 }

# UUID Remote
Function Get-UUIDTarget{
    param($RemoteHost)
    $UUID = ""
    try{
        $UUID = (Get-WmiObject -Class Win32_ComputerSystemProduct -ComputerName $RemoteHost).UUID
    }Catch{
        try{
            $UUID = (Get-CimInstance -Class Win32_ComputerSystemProduct -ComputerName $RemoteHost).UUID
        }catch{}
    }
    Return $UUID
}

#Function StringCleanup
Function Edit-StringRUC {
    param($String)
    $String = $String -replace ''
    $String = $String -replace '\s',''
    $String = $String -replace ':',''
    $String = $String -replace ';',''
    $String = $String -replace '\\','_'
    $String = $String -replace '/',''
    $String = $String -replace '-',''
    $String = $String -replace '~',''
    $String = $String -replace '%',''
    $String = $String -replace '$',''
    $String = $String -replace '\*',''
    $String = $String -replace '^',''
    $String = $String -replace '@',''
    $String = $String -replace '!',''
    $String = $String -replace '#',''
    $String = $String -replace '¤',''
    $String = $String -replace '&',''
    $String = $String -replace '{',''
    $String = $String -replace '}',''
    $String = $String -replace '\[',''
    $String = $String -replace ']',''
    $String = $String -replace '\(',''
    $String = $String -replace '\)',''
    $String = $String -replace '=',''
    $String = $String -replace '\?',''
    $String = $String -replace '¨',''
    $String = $String -replace '\+',''
    return $String
}

#TimeStamp
Write-Host "Creating Time Stamp"
$Stamp_Date = Get-Date -Format "yyyyMMdd"
$Stamp_Date
$Stamp_Time = Get-Date -Format "HHmmss"
$Stamp_Time

# Misc
Write-Host "Setting variables"
$Host_Name = (Get-CimInstance -ClassName Win32_ComputerSystem).Name
$HW_Model = (Get-CimInstance -ClassName Win32_ComputerSystem).Model
$HW_Model = $HW_Model -replace '\s', '_'

$Array_DriveList = Get-PSDrive -PSProvider FileSystem
$String_DriveList = ""
$Letter_ThisUSBDrive = ""

$List_FileCopy = @()
$ID_NewAudit = $Stamp_Date + "-" + $Stamp_Time + "-" + $HW_Serialnumber + "-" + $HW_Model
$Name_SaveFolder = '_LogFiles\FailedOSD\' + $Stamp_Date + "\" + $ID_NewAudit

# Drive Letters in a readable format
$Array_DriveList | ForEach-Object{
    $String_DriveList = $String_DriveList + $_
}

# Identify USB
Write-Host "Finding USBDrive"
For($i = 0; $i -lt $String_DriveList.Length; $i++){
    $HardDrive = $String_DriveList[$i] + ":\"
    IF(Test-Path -Path ($HardDrive + $Filename_ThisScript)){
        $Letter_ThisUSBDrive = $HardDrive
        Write-Host $Letter_ThisUSBDrive "USB"
    }
}
# Store files at this path
$Path_NewFolder = ($Letter_ThisUSBDrive + $Name_SaveFolder)

# Show findings
For($i = 0; $i -lt $String_DriveList.Length; $i++){
    $HardDrive = $String_DriveList[$i] + ":\"
    IF(!(Test-Path -Path ($HardDrive + $Filename_ThisScript))){
        Write-Host $HardDrive
        For($_i = 0; $_i -lt $List_FileSearch.Count; $_i++){
            Write-Host "Finding " + $List_FileSearch[$_i]
            $File_AddToList = (Get-ChildItem -Path $HardDrive -Filter $List_FileSearch[$_i] -Recurse -ErrorAction SilentlyContinue -Force | ForEach-Object{$_.FullName})
            Write-Host $File_AddToList
            $File_AddToList | ForEach-Object {$List_FileCopy += $File_AddToList}
        }
    }
}

$List_FileCopy = $List_FileCopy | Sort-Object -Unique
Write-Host $List_FileCopy

# Create Logfile
$Path_AuditLogFile = $Path_NewFolder + '\' + $HW_Serialnumber + "_ipconfig_all_and_ping_test.txt"
New-Item $Path_AuditLogFile -Force

# Add network info to Logfile
$Info_Net = ipconfig /all
$Info_Net | Add-Content -Path $Path_AuditLogFile -Encoding Unicode
Add-Content "" -Path $Path_AuditLogFile -Encoding Unicode
Add-Content "################" -Path $Path_AuditLogFile -Encoding Unicode
Add-Content "" -Path $Path_AuditLogFile -Encoding Unicode

# Test connection and save to logfile
For($i = 0; $i -lt $List_Ping.Count; $i++){
    $New_Ping = ""
    IF($List_Ping[$i].Contains(".")){
        try{
            Write-Host "Ping: " $List_Ping[$i].ToString()
            $New_Ping = Test-Connection $List_Ping[$i] -ErrorAction Stop
            Write-Host $New_Ping
            Add-Content $New_Ping -Path $Path_AuditLogFile -Encoding Unicode
            Add-Content "" -Path $Path_AuditLogFile -Encoding Unicode
        }
        catch{
            Write-Warning "Warning: " + $_
        }
    }
}
Add-Content "" -Path $Path_AuditLogFile -Encoding Unicode

# Copy files to USB
New-Item -ItemType Directory -Force -Path $Path_NewFolder
$List_FileCopy | ForEach-Object {
    $This_Row = $_.ToString()
    $This_Row = Edit-StringRUC $This_Row
    $File_NewName = ('\' + $ID_NewAudit + "-" + $This_Row)
    $File_SaveNewCopy = ($Path_NewFolder + $File_NewName)
    Write-Host "Copying " $_ "to" $File_SaveNewCopy
    Copy-Item -Path $_ -Destination $File_SaveNewCopy
}

# Copy Windows\Panther folder if exists
IF(Test-Path -Path C:\Windows\panther){
    Copy-Item -Path C:\Windows\panther -Destination ($Path_NewFolder + '\Windows\panther\') -Recurse -Force
}
# Copy Windows\CCM\Logs folder if exists
IF(Test-Path -Path C:\Windows\CCM\Logs){
    Copy-Item -Path C:\Windows\CCM\Logs -Destination ($Path_NewFolder + '\Windows\CCM\') -Recurse -Force
}
# Copy Windows\ccmsetup\Logs folder if exists
IF(Test-Path -Path C:\Windows\ccmsetup\Logs){
    Copy-Item -Path C:\Windows\ccmsetup\Logs -Destination ($Path_NewFolder + '\Windows\ccmsetup\') -Recurse -Force
}

# Get System event log
$Path_NewFile = $Path_NewFolder + "\" + $HW_Serialnumber + "_System_Log.evtx"
wevtutil epl System $Path_NewFile

# Get Applications event log
$Path_NewFile = $Path_NewFolder + "\" + $HW_Serialnumber + "_Application_Log.evtx"
wevtutil epl Application $Path_NewFile

# Collect info about Installed Drivers
$Path_NewFile = $Path_NewFolder + "\" + $HW_Serialnumber + "_Drivers.txt"
Get-WindowsDriver -Online -All |  Out-File -FilePath $Path_NewFile

# Collect UUID
$Path_NewFile = $Path_NewFolder + "\" + $HW_Serialnumber + "_UUID.txt"
Get-UUIDLocal | Out-File -FilePath $Path_NewFile

# Collect Certification Info (ALL)
$Path_NewFile = $Path_NewFolder + "\" + $HW_Serialnumber + "_Cert_ALL.txt"
Get-ChildItem Cert:\ -Recurse | Out-File -FilePath $Path_NewFile

# Collect Certification Info ("LocalMachine\My")
$Path_NewFile = $Path_NewFolder + "\" + $HW_Serialnumber + "_Cert_LocalMachine_My.txt"
Get-ChildItem -path cert:\LocalMachine\My | Format-List Name, Subject, FriendlyName, DnsNameList, EnhancedKeyUsageList, SendAsTrustedIssuer | Out-File -FilePath $Path_NewFile

$Path_BIOSConf = (Get-ChildItem -Path $Letter_ThisUSBDrive -Filter "BiosConfigUtility64.exe" -Recurse -ErrorAction SilentlyContinue -Force | ForEach-Object{$_.FullName})
if(Test-Path($Path_BIOSConf)){
    $Path_NewFile = $Path_NewFolder + "\" + $HW_Serialnumber + "_BIOSConf.txt"
    . $Path_BIOSConf /get:$Path_NewFile
}



# End of script
Write-Host
Write-Host
Write-Host "DONE!"
Write-Host Exit
Exit

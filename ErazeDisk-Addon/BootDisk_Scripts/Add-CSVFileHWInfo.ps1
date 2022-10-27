######################## Custom Variables ##########################
$Foldername_Scripts = 'BootDisk_Scripts\'
$Foldername_SaveLogs = '_LogFiles\CSV_HW\'
$Filename_ThisScript = $Foldername_Scripts + "Add-CSVFileHWInfo.ps1"
$Filename_UtilityModule = 'Module-MiscFunctions.psm1'
####################################################################
Write-Host 'Start Script'
# Clear-Host
Function Set-WindowTitle{
    param($Title)
    $host.UI.RawUI.WindowTitle = $Title
}

Write-Host 'Identify Drives'
$String_DriveList = ""
Get-PSDrive -PSProvider FileSystem | ForEach-Object{
    $String_DriveList = $String_DriveList + $_
}

Write-Host 'Identify USB Drive'
$Drive_ThisUSB = ""
For($i = 0; $i -lt $String_DriveList.Length; $i++){
    $HardDrive = $String_DriveList[$i] + ":\"
    IF(Test-Path -Path ($HardDrive + $Filename_ThisScript)){
        $Drive_ThisUSB = $HardDrive
        Set-WindowTitle -Title $Drive_ThisUSB "USB"
    }
}
$Path_ScriptFolder = $Drive_ThisUSB + $Foldername_Scripts

Write-Host 'Importing Utility Module'
$Path_UtilityModule = $Path_ScriptFolder + $Filename_UtilityModule
IF([System.IO.File]::Exists($Path_UtilityModule)){
    Import-Module $Path_UtilityModule -Verbose
}

Write-Host 'Update .CSV File'
$CSVFirstLine = 'ComputerSerialNumber,SMBIOSUUID,SystemManufacturer,SystemModel,ComputerName,Processor,RAM,DiskSpace'
$ComputerSerialNumber = (Get-ComputerSystemSerialnumber)
$SMBIOSUUID = (Get-ComputerSystemUUID)
$SystemManufacturer = (Get-ComputerSystemManufacturer)
$SystemModel = (Get-ComputerSystemModel)
$ComputerName = (Get-ComputerSystemName)
$Processor = (Get-ComputerSystemProcessorName)
$RAM = (Get-TotalRAMCapacity)
$DiskSpace = (Get-TotalHardDriveCapacity)

$NewContent = "$ComputerSerialNumber,$SMBIOSUUID,$SystemManufacturer,$SystemModel,$ComputerName,$Processor,$RAM,$DiskSpace"
$CSVFilePath = $Drive_ThisUSB + $Foldername_SaveLogs + (Get-Date -Format "yyyyMMdd") + ".csv"
Update-CSVFile -CSVFilePath $CSVFilePath -CSVFirstLine $CSVFirstLine -NewContent $NewContent

Write-Host 'Hardware information file Update Complete!'
Set-WindowTitle -Title 'Hardware information file Update Complete!'
########################
#Start-Sleep -Seconds 5
########################
######################## Custom Variables ##########################
$Foldername_Scripts = 'BootDisk_Scripts\'
#$Foldername_SaveLogs = '_LogFiles\CSV_ResetBIOS\'
$Filename_ThisScript = $Foldername_Scripts + "Set-BIOSFactorySettings.ps1"
$Filename_UtilityModule = 'Module-MiscFunctions.psm1'
####################################################################

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

Set-WindowTitle -Title 'BIOS Defaults!'
# HP BIOS only

if((Get-ComputerSystemManufacturer) -eq "HP"){
    $Path_BIOSConf = (Get-ChildItem -Path $Drive_ThisUSB -Filter "BiosConfigUtility64.exe" -Recurse -Force | ForEach-Object{$_.FullName})
    if(Test-Path($Path_BIOSConf)){
        $Path_BIOS_PW = (Get-ChildItem -Path $Drive_ThisUSB -Filter "*HP_BIOS_PW.bin" -Recurse -Force | ForEach-Object{$_.FullName})
        $Path_BIOS_PW | ForEach-Object {
            if(Test-Path($_)){
                . $Path_BIOSConf /cpwdfile:$_ /nspwdfile:""
            }
        }
        . $Path_BIOSConf /setdefaults
        Set-WindowTitle -Title 'HP BIOS Reset!'
        Write-Host 'HP BIOS Reset!'
    }
}
else{
    Set-WindowTitle -Title 'BIOS Reset failed :('
    Write-Host 'BIOS Reset failed :('
}

#######################
Start-Sleep -Seconds 10
#######################
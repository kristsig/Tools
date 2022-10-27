## Console and Window
Function Set-WindowTitle{
    param($Title)
    $host.UI.RawUI.WindowTitle = $Title
}

## Filesystem
Function Get-DriveContainingFile{
    Param($File)
    $Drives = ""
    $USBDrive = ""
    Get-PSDrive -PSProvider FileSystem | ForEach-Object{
        $Drives += $_
    }
    For($i = 0; $i -lt $Drives.Length; $i++){
        $Drive = $Drives[$i] + ":\"
        IF([System.IO.File]::Exists($Drive + $File)){
            $USBDrive = $Drive
        }
    }
    Return $USBDrive
}

Function Update-CSVFile{
    param($CSVFilePath, $CSVFirstLine, $NewContent)
    if(!(Test-Path ($CSVFilePath))){
        New-Item -Path $CSVFilePath -Force
        Add-Content -Path $CSVFilePath -Value $CSVFirstLine
    }
    Add-Content -Path $CSVFilePath -Value $NewContent
}

## Text
Function Edit-StringReplaceItems {
    param($String, $Replace, $Insert)
    IF(($String.Length -gt 0) -And ($Replace.Length -gt 0) -And ($Insert.Length -gt 0)){
        foreach ($char in $Replace) {
            $String = $String -replace $char,$Insert            
        }
    }
    return $String
}

Function Edit-StringRemoveItems {
    param($String, $Remove)
    IF(($String.Length -gt 0) -And ($Remove.Length -gt 0)){
        foreach ($char in $Remove) {
            $String = $String -replace $char,''
        }
    }
    return $String
}

Function Start-TextToSpeach{
    param($Text)
    Add-Type -AssemblyName System.speech
    $SpeakText = New-Object System.Speech.Synthesis.SpeechSynthesizer
    $SpeakText.Speak($Text)
}

## Hardware Information
Function Get-ComputerSystemUUID{
    $UUID = ""
    try{
        $UUID = (Get-WmiObject -Class Win32_ComputerSystemProduct).UUID
    }Catch{
        try{
            $UUID = (Get-CimInstance -Class Win32_ComputerSystemProduct).UUID
        }catch{$UUID = 'Error'}
    }
    Return $UUID
}

Function Get-ComputerSystemUUIDRemote{
    param($HostName)
    $UUID = ""
    try{
        $UUID = (Get-WmiObject -Class Win32_ComputerSystemProduct -ComputerName $HostName).UUID
    }catch{
        try{
            $UUID = (Get-CimInstance -Class Win32_ComputerSystemProduct -ComputerName $HostName).UUID
        }catch{$UUID = 'Error'}
    }
    Return $UUID
}

Function Get-ComputerSystemSerialnumber{
    try{
        $Serialnumber = (Get-WmiObject -ClassName Win32_ComputerSystemProduct).IdentifyingNumber
    }catch{
        try{
            $Serialnumber = (Get-CimInstance -ClassName Win32_ComputerSystemProduct).IdentifyingNumber
        }catch{$Serialnumber = 'Error'}
    }
    Return $Serialnumber
}

Function Get-ComputerSystemSerialnumberRemote{
    param($HostName)
    try{
        $Serialnumber = (Get-CimInstance -ClassName Win32_ComputerSystemProduct -ComputerName $HostName).IdentifyingNumber
    }catch{
        try{
            $Serialnumber = (Get-CimInstance -ClassName Win32_ComputerSystemProduct -ComputerName $HostName).IdentifyingNumber
        }catch{$Serialnumber = 'Error'}
    }
    Return $Serialnumber
}

Function Get-ComputerSystemName{
    try{
        $Name = (Get-WmiObject -ClassName Win32_ComputerSystem).Name
    }catch{
        try{
            $Name = (Get-CimInstance -ClassName Win32_ComputerSystem).Name
        }catch{$Name = 'Error'}
    }
    Return $Name
}

Function Get-ComputerSystemModel{
    try{
        $Model = (Get-WmiObject -ClassName Win32_ComputerSystem).Model
    }catch{
        try{
            $Model = (Get-CimInstance -ClassName Win32_ComputerSystem).Model
        }catch{$Model = 'Error'}
    }
    Return $Model
}

Function Get-ComputerSystemManufacturer{
    try{
        $Manufacturer = (Get-WmiObject -ClassName Win32_ComputerSystem).Manufacturer
    }catch{
        try{
            $Manufacturer = (Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer
        }catch{$Manufacturer = 'Error'}
    }
    Return $Manufacturer
}

Function Get-ComputerSystemVendor{
    try{
        $Vendor = (Get-WmiObject -ClassName Win32_ComputerSystem).Vendor
    }catch{
        try{
            $Vendor = (Get-CimInstance -ClassName Win32_ComputerSystem).Vendor
        }catch{$Vendor = 'Error'}
    }
    Return $Vendor
}

Function Get-ComputerSystemRAMInfo{
    try{
        $RAMInfo = (Get-WmiObject -ClassName Win32_Physicalmemory)
    }catch{
        try{
            $RAMInfo = (Get-CimInstance -ClassName Win32_Physicalmemory)
        }catch{$RAMInfo = 'Error'}
    }
    Return $RAMInfo
}

Function Get-TotalRAMCapacity{
    try{
        $RAMTotalCapacity = 0
        $RAMInformation = Get-ComputerSystemRAMInfo
        for($i = 0;$i -lt $RAMInformation.Length; $i++){
            $RAMTotalCapacity += $RAMInformation[$i].Capacity
        }
    }catch{$RAMTotalCapacity = 'Error'}
    Return $RAMTotalCapacity
}

Function Get-ComputerSystemHardriveInfo{
    try{
        $HardriveInfo = (Get-WmiObject -ClassName win32_logicaldisk)
    }catch{
        try{
            $HardriveInfo = (Get-CimInstance -ClassName win32_logicaldisk)
        }catch{$HardriveInfo = 'Error'}
    }
    Return $HardriveInfo
}

Function Get-TotalHardDriveCapacity{
    $HDDotalCapacity = 0
    try{
        $NoList = (Get-DriveContainingFile -File 'BootDisk_Scripts\Module-MiscFunctions.psm1')
        $HDDInformation = Get-ComputerSystemHardriveInfo
        for($i = 0;$i -lt $HDDInformation.Length; $i++){
            if(!(($HDDInformation[$i].DeviceID + '\') -eq $NoList)){
                $HDDotalCapacity += $HDDInformation[$i].Size
            }
        }
    }catch{$HDDotalCapacity = 'Error'}
    Return $HDDotalCapacity
}

Function Get-ComputerSystemProcessorName{
    try{
        $ProcName = (Get-WmiObject -ClassName Win32_processor).Name
    }catch{
        try{
            $ProcName = (Get-CimInstance -ClassName Win32_processor).Name
        }catch{$ProcName = 'Error'}
    }
    Return $ProcName
}

## Time
Function New-DateStamp{
    $DateStamp = Get-Date -Format "yyyyMMdd"
    return $DateStamp
}

Function New-TimeStamp{
    $TimeStamp = Get-Date -Format "HHmmss"
    return $TimeStamp
}

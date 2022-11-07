## A few customizable settings
$Timestamp = Get-Date -Format "yyyyMMdd"
$TargetDNS = "" # This setting can help if target is on a different subnet or script runs via VPN.
$TargetHostname_MinLength = 10
$Filename_ScanLog = "ScanLog-" + $Timestamp+ ".txt"
$Path_ScanLog = 'C:\temp\scan\' + $Filename_ScanLog

#####
IF(!(Test-Path($Path_ScanLog))){
    New-Item -Path $Path_ScanLog -Force
}

Function Start-TextToSpeach{
    param($Text)
    Add-Type -AssemblyName System.speech
    $SpeakText = New-Object System.Speech.Synthesis.SpeechSynthesizer
    Start-Process Powershell {$SpeakText.Speak($Text)} -ErrorAction SilentlyContinue
}

Clear-Host
Function Invoke-ShutdownCommand{
    param ($TargetHost, $TargetDNS, $CommandShutdownMethod)
    $TargetID = $TargetHost + $TargetDNS
    $StopMessage = "Stopping: "
    If($TargetHost.length -le 13){
        $ZerothCommand = ("""shutdown /s /t 2 /m \\$TargetHost""")
        $FirstCommand = ("""shutdown /s /t 2 /m \\$TargetID""")
        $SecondCommand = ("""Stop-Computer -ComputerName $TargetHost""")
        $ThirdCommand = ("""Stop-Computer $TargetID""")
        If($StopMethod -eq "R"){
            $ZerothCommand = ("""shutdown /r /t 2 /m \\$TargetHost""")
            $FirstCommand = ("""shutdown /r /t 2 /m \\$TargetID""")
            $SecondCommand = ("""Restart-Computer $TargetHost""")
            $ThirdCommand = ("""Restart-Computer $TargetID""")
            $StopMessage = "Rebooting: "
        }
        Write-Host $StopMessage $TargetID
        try{
            Start-Process -WindowStyle Hidden powershell.exe -ArgumentList $ZerothCommand
        }Catch{}
        try{
            Start-Process -WindowStyle Hidden powershell.exe -ArgumentList $FirstCommand
        }Catch{}
        try{
            Start-Process -WindowStyle Hidden powershell.exe -ArgumentList $SecondCommand
        }Catch{}
        try{
            Start-Process -WindowStyle Hidden powershell.exe -ArgumentList $ThirdCommand
        }Catch{}
   }
   Else{Start-TextToSpeach -Text $TargetHost}
}

Function Start-ScanToStopComputer{
    param($TargetID_MinLength, $TargetDNS, $LogFile)
    $NewScan = ""
    $CommandShutdownMethod = "S"
    $ContinueLoop = $true
    $Name_ThisComputer = (Get-CimInstance -ClassName Win32_ComputerSystem).Name
    While($ContinueLoop){
        Write-Host ""
        Write-Host @("(S)hutdown/(R)estart/(Q)uit
Enter target host ID:")
        $NewScan = Read-Host
        If($NewScan -match ','){
            $SplitString = $NewScan -split ','
            $SplitString | ForEach-Object{
                If($_ -match '#'){
                    Write-Host 'Found #'
                    $_ = ''
                }
                If($_ -match ' '){
                    Write-Host 'Found " "'
                    $_ = ''
                }
                If($_.length -ge 10){
                    Write-Host $NewScan ' Set To '$_
                    $NewScan = $_
                }
            }
        }
        If($NewScan -match '\.'){
            Write-Host "Dot!?"
            Start-TextToSpeach -Text $NewScan
            $NewScan = ''
        }
        if($NewScan -eq $Name_ThisComputer){
            $NewScan = ""
        }
        If($NewScan.Length -gt 0){
            If(($NewScan.StartsWith('S')) -or ($NewScan.StartsWith('s'))){
                Write-Host "Set to Shutdown:"
                $CommandShutdownMethod = "S"
            }
            If(($NewScan.StartsWith('R')) -or ($NewScan.StartsWith('r'))){
                Write-Host "Set to Reboot:"
                $CommandShutdownMethod = "R"
            }
            if($NewScan.length -ge $TargetID_MinLength){
                Invoke-ShutdownCommand -TargetHost $NewScan -TargetDNS $TargetDNS -CommandShutdownMethod $CommandShutdownMethod
                Add-Content -Value $NewScan -Path $LogFile -Force
            }
            If(($NewScan.StartsWith('Q')) -or ($NewScan.StartsWith('q'))){
                $ContinueLoop = $false
            }
        }
        Else{
            Write-Host "String to short?"
        }
    }
    Write-Host "Exit:"
}

Start-ScanToStopComputer -TargetID_MinLength $TargetHostname_MinLength -TargetDNS $TargetDNS -LogFile $Path_ScanLog

Exit
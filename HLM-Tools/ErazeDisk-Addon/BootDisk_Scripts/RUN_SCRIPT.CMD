@ECHO OFF

TITLE Initializing Network...
wpeutil.exe InitializeNetwork
cls

TITLE Preparing K-Disk...
FOR %%i IN (C D E F G H I J K L M N O P Q R S T U V W X Y Z) DO IF EXIST %%i:\BootDisk_Scripts\RUN_SCRIPT.CMD ( 
SET DrvLtr=%%i
GOTO END)
:END
ECHO BootDisk is: %DrvLtr%:

TITLE Run Powershell scripts...
Powershell Set-ExecutionPolicy Unrestricted
Powershell %DrvLtr%:\BootDisk_Scripts\Add-CSVFileOSD.ps1
Powershell %DrvLtr%:\BootDisk_Scripts\Set-BIOSFactorySettings.ps1
FOR /F "usebackq" %%i IN (`hostname`) DO SET CURRENTHOST=%%i
ECHO %CURRENTHOST%
SET TIMEVAR=%TIME::=%
SET TIMEVAR=%TIMEVAR:,=%

IF EXIST "C:\CopySMSTSLog.bat" (
  SET USBSTICK="C:\Logs\"
)
IF EXIST "D:\CopySMSTSLog.bat" (
  SET USBSTICK="D:\Logs\"
)
IF EXIST "E:\CopySMSTSLog.bat" (
  SET USBSTICK="E:\Logs\"
)
IF EXIST "F:\CopySMSTSLog.bat" (
  SET USBSTICK="F:\Logs\"
)
IF EXIST "G:\CopySMSTSLog.bat" (
  SET USBSTICK="G:\Logs\"
)
IF EXIST "H:\CopySMSTSLog.bat" (
  SET USBSTICK="H:\Logs\"
)
IF EXIST "I:\CopySMSTSLog.bat" (
  SET USBSTICK="I:\Logs\"
)


IF EXIST "x:\windows\temp\smstslog\smsts.log" (
  SET LogLocation="x:\windows\temp\smstslog\smsts.log"
)
IF EXIST "x:\windows\temp\smstslog\smsts.log" (
  SET LogLocation="x:\windows\temp\smstslog\smsts.log"
)
IF EXIST "x:\smstslog\smsts.log" (
  SET LogLocation="x:\smstslog\smsts.log"
)
IF EXIST "C:\_SMSTaskSequence\Logs\Smstslog\smsts.log" (
  SET LogLocation="C:\_SMSTaskSequence\Logs\Smstslog\smsts.log"
)
IF EXIST "c:\_SMSTaskSequence\Logs\Smstslog\smsts.log" (
  SET LogLocation="c:\_SMSTaskSequence\Logs\Smstslog\smsts.log"
)
IF EXIST "c:\windows\system32\ccm\logs\Smstslog\smsts.log" (
  SET LogLocation="c:\windows\system32\ccm\logs\Smstslog\smsts.log"
)
IF EXIST "c:\windows\sysWOW64\ccm\logs\Smstslog\smsts.log" (
  SET LogLocation="c:\windows\sysWOW64\ccm\logs\Smstslog\smsts.log"
)
IF EXIST "c:\windows\system32\ccm\logs\smsts.log" (
  SET LogLocation="c:\windows\system32\ccm\logs\smsts.log"
)
IF EXIST "c:\windows\sysWOW64\ccm\logs\smsts.log" (
  SET LogLocation="c:\windows\sysWOW64\ccm\logs\smsts.log"
)
IF EXIST "c:\windows\ccm\logs\smstslog\smsts.log" (
  SET LogLocation="c:\windows\ccm\logs\smstslog\smsts.log"
)
IF EXIST "c:\windows\ccm\logs\smsts.log" (
  SET LogLocation="c:\windows\ccm\logs\smsts.log"
)


xcopy %LogLocation% %USBSTICK%%DATE%-%TIMEVAR%\%CURRENTHOST%\smsts.log
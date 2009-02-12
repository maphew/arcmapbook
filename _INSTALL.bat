REM register component
%windir%\system32\regsvr32.exe /s ".\Visual_Basic\NWMapBookPrj.dll"
%windir%\system32\regsvr32.exe /s ".\Visual_Basic\NWMapBookUIPrj.dll"

REM register components in appropriate component categories
.\NWMapBookUIPrj.reg

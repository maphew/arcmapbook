@echo off
echo.

echo.	De-registering mapbook DLLs...
regsvr32 /s /u ".\Visual_Basic\DSMapBookUIPrj.dll"
regsvr32 /s /u ".\Visual_Basic\DSMapBookPrj.dll"

echo.	Removing mapbook registry keys...
regedit /s ".\Visual_Basic\de-register_component_category.reg"

echo.
pause
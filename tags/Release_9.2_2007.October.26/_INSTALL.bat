@echo off
echo.

echo.	Registering mapbook DLLs...
regsvr32 /s ".\Visual_Basic\DSMapBookPrj.dll"
regsvr32 /s ".\Visual_Basic\DSMapBookUIPrj.dll"

echo.	Adding mapbook registry keys...
regedit /s ".\Visual_Basic\register_component_category.reg"

echo.
pause
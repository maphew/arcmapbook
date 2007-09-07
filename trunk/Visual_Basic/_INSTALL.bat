@echo off
echo.

echo.	Registering mapbook DLLs...
regsvr32 /s ".\DSMapBookPrj.dll"
regsvr32 /s ".\DSMapBookUIPrj.dll"

echo.	Adding mapbook registry keys...
regedit /s .\register_component_category.reg

echo.
pause
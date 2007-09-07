@echo off
echo.

echo.	De-registering mapbook DLLs...
regsvr32 /s /u ".\DSMapBookUIPrj.dll"
regsvr32 /s /u ".\DSMapBookPrj.dll"

echo.	Removing mapbook registry keys...
regedit /s .\de-register_component_category.reg

echo.
pause
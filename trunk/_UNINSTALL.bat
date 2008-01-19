@echo off
echo.
echo.	De-registering mapbook DLLs...
for /r %%a in (DSMapBook*.dll) do (
	echo.		%%a
	regsvr32 /s /u "%%a"
	)

echo.
echo.	Removing mapbook registry keys...
for /r %%a in (de-register_*.reg) do (
	echo.		%%a
	regedit /s "%%a"
	)

echo.
pause
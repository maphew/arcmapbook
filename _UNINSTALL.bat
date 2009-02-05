@echo off
setlocal
if /i [%1]==[debug] (echo *** Running in debug mode) else (
   set _opt=/s
   echo.
   echo. If there are problems try:
   echo.
   echo.    %0 debug
   echo.
   )
   
echo.
echo --- De-registering mapbook DLLs...
for /r %%a in (DSMapBook*.dll) do (
	echo  %%a
   %WINDIR%\system32\regsvr32 %_opt% /u "%%a"
	)

echo.
echo --- Removing mapbook registry keys...
for /r %%a in (de-register_*.reg) do (
	echo  %%a
	%WINDIR%\regedit %_opt% "%%a"
	)

echo.
pause
endlocal

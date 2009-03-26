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
echo +++ Registering mapbook DLLs...
%WINDIR%\system32\regsvr32 %_opt% ".\Visual_Basic\NWMapBookPrj.dll"
%WINDIR%\system32\regsvr32 %_opt% ".\Visual_Basic\NWMapBookUIPrj.dll"

echo +++ Adding mapbook registry keys...
%WINDIR%\regedit %_opt% ".\Visual_Basic\NWMapBookUIPrj.reg"

echo.
pause
endlocal

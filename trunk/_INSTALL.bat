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
%WINDIR%\system32\regsvr32 %_opt% ".\Visual_Basic\DSMapBookPrj.dll"
%WINDIR%\system32\regsvr32 %_opt% ".\Visual_Basic\DSMapBookUIPrj.dll"

echo +++ Adding mapbook registry keys...
if DEFINED ProgramFiles(x86) (
   %WINDIR%\regedit %_opt% ".\Visual_Basic\register_component_category_x64.reg"
   ) else (
   %WINDIR%\regedit %_opt% ".\Visual_Basic\register_component_category.reg"
   )

echo.
pause
endlocal

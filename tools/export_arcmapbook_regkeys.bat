@echo off
echo.
echo. Arcmapbook Debug Registry -- exporting registry keys for %username%
echo.

:: Remove old export
if exist %username%_arcmapbook_registry.txt del %username%_arcmapbook_registry.txt

:: Normal entries
for %%g in (AC7622A7-6D66-4D2B-9AE0-EB70BD262B53 E918E787-8B4E-4D51-877C-AD67905C6109 B121B1BA-5420-464B-802A-7A6C89123093 DC395506-3391-4207-99D4-C70851BAE9EA 1DA56C9C-4646-41B8-93CE-61AB6F04D982 122B316F-67A6-42D4-B76D-63BFB6210393 BBAF9983-58D2-40D7-A093-FE564EA8966E) do (
   regedit /e "%temp%\%username%_%%g.txt" "HKEY_CLASSES_ROOT\CLSID\{%%g}"
   type "%temp%\%username%_%%g.txt" >> %username%_arcmapbook_registry.txt
   del "%temp%\%username%_%%g.txt"
   )

:: 64bit XP entries   
for %%g in (AC7622A7-6D66-4D2B-9AE0-EB70BD262B53 E918E787-8B4E-4D51-877C-AD67905C6109) do (
   regedit /e "%temp%\%username%_%%g.txt" "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Wow6432Node\CLSID\{%%g}"
   type "%temp%\%username%_%%g.txt" >> %username%_arcmapbook_registry.txt
   del "%temp%\%username%_%%g.txt"
   )

echo.
echo. Results are in %username%_arcmapbook_registry.txt
echo.

pause
@echo off
pushd %~dp0
set outdir=D:\code\mapbook_extras\release_test\%1

echo %outdir%
if exist %outdir% goto :outExists

call :Main
Call :Extras

start %outdir%
goto :eof

:: -------------------------------------------------------------------- 
:Main
  :: Copy single, named files from manifest to release folder
  mkdir %outdir%\Visual_Basic
	for /f "eol=¬ delims=" %%g in (manifest.txt) do (
		copy /v "%%g" "%outdir%\%%g"
		)
	goto :eof

:Extras
  :: Recursively copy directories & contents in manifest to release folder
	for /f "eol=¬ delims=" %%g in (manifest-extras.txt) do (
		xcopy /s/i/v/q "%%g" "%outdir%\%%g"
		)
  goto :eof
  
:outExists
	echo.
	echo.	%outdir% already exists.
	goto :eof

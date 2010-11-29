@echo off
pushd %~dp0
set outdir=D:\code\mapbook_extras\release_test\%1

echo %outdir%
if exist %outdir% goto :outExists

mkdir %outdir%
mkdir %outdir%\Visual_Basic

:Main
	for /f "eol=¬ delims=" %%g in (manifest.txt) do (
		copy "%%g" "%outdir%\%%g"
		)
	start %outdir%
	goto :eof


:outExists
	echo.
	echo.	%outdir% already exists.
	goto :eof

@echo off
echo Building ExcelTrans:
PowerShell -Command ".\psake.ps1" "%1"

If Not "%NugetPackagesDir%" == "" xcopy .\_build\*.nupkg %NugetPackagesDir% /Y/Q
If Not "%NugetPackagesDir%" == "" del %NugetPackagesDir%\*.symbols.nupkg /Q

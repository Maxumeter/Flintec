@echo off
rem Baut Stundennachweis.exe neu - mit dem C#-Compiler, der bei Windows dabei ist.
"%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /nologo /target:winexe /platform:anycpu /optimize+ /win32icon:"%~dp0Stundennachweis.ico" /out:"%~dp0..\Stundennachweis.exe" "%~dp0Starter.cs"
pause

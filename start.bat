@echo off

PATH=%PATH%;%~d0%~p0lib

.\jre\bin\java.exe  -jar excelsikuli.jar sikuli_IF.xlsx 
pause

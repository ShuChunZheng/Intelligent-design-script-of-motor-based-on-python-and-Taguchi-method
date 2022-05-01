@echo off
echo get admin rights
cacls.exe "%SystemDrive%\System Volume Information" >nul 2>nul
if %errorlevel%==0 goto isAdmin

if exist "%temp%\getAdmin.vbs" del /f /q "%temp%\getAdmin.vbs"
echo Set shell = CreateObject^("Shell.Application"^)>"%temp%\getAdmin.vbs"
echo shell.ShellExecute "%~s0","","","runas",1 >>"%temp%\getAdmin.vbs"
echo WScript.Quit >>"%temp%\getAdmin.vbs"
"%temp%\getAdmin.vbs" /f
if exist "%temp%\getAdmin.vbs" del /f /q "%temp%\getAdmin.vbs"
exit

:isAdmin


python -OO -m compileall C:\Users\SJ\PycharmProjects\pythonProject1
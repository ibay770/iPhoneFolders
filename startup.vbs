On Error Resume Next
set WshShell = WScript.CreateObject("WScript.Shell")
Set WshProcEnv = WshShell.Environment("Process")
process_architecture= WshProcEnv("PROCESSOR_ARCHITECTURE")
system_architecture= WshProcEnv("PROCESSOR_ARCHITEW6432")
if system_architecture="" then system_architecture=process_architecture

strOS = WshProcEnv("OS")
If strOS = "Windows_NT" Then
   strVerKey = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\"
   Version = WshShell.regread(strVerKey & "CurrentVersion")
end if

if (Version = "6.1" and system_architecture<>"x86") then
    MsgBox "Sorry, Windows7 x64 is not supported"
    WScript.Quit
end if

if (system_architecture="x86") then
    WshShell.Exec "%windir%\explorer.exe /e,::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\::{D9D587F5-8284-45cc-AA5C-D2123D8852D9}"
else
    WshShell.Exec "%Systemroot%\SysWOW64\explorer.exe /separate /e,::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\::{D9D587F5-8284-45cc-AA5C-D2123D8852D9}"
end if

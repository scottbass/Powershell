@echo off

if [%1]==[] (
  echo Missing Excel XML file
  exit
) else (
  set xlfile=%1
)  

if [%2]==[] (
  set xlformat=xlsb
) else (
  set xlformat=%2
)  

C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe -noprofile -executionpolicy unrestricted -windowstyle minimized -command "E:\Powershell\Scripts\ShrinkExcel.ps1 -xlfile '%xlfile%' -xlformat %xlformat%" 2>&1

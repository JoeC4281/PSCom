Dim obj
Dim result

' Create the COM object
Set obj = CreateObject("PSCom.PSScript")

' Test ExecuteScript with a sample PowerShell script path
result = obj.ExecuteScript("E:\Utils\WinVer.ps1")
WScript.Echo result

' Test ExecuteCommand with a sample PowerShell command(s)
result = obj.ExecuteCommand("Get-Date; 2025-1959")
WScript.Echo result

Set obj = Nothing

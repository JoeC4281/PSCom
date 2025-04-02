Dim obj
Dim result

' Create the COM object
Set obj = CreateObject("PSCom.PSScript")

' Test Add2 method
result = obj.Add2(5) ' This should return 7
WScript.Echo "Add2 Result: " & result

' Test ExecuteScript with a sample PowerShell script path
result = obj.ExecuteScript("E:\Utils\ballon.ps1")
result = obj.ExecuteScript("E:\Utils\WinVer.ps1")
WScript.Echo "ExecuteScript Result: " & result

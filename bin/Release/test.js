var obj, result;

// Create the COM object
obj = new ActiveXObject("PSCom.PSScript");

// Test ExecuteScript with a sample PowerShell script path
result = obj.ExecuteScript("E:\\Utils\\WinVer.ps1");
WScript.Echo(result);

// Test ExecuteCommand with a sample PowerShell command(s)
result = obj.ExecuteCommand("Get-Date; 2025-1959");
WScript.Echo(result);

obj = null;

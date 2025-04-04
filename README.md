# PSCom
This is a 64-bit COM ActiveX .dll that allows me to run a PowerShell Script,\
or a PowerShell Command(s),\
from 64-bit or 32-bit VBScript.

You will need to use <span style="font-family: Courier New; font-size: 20px;">RegASM.exe</span> to register the <span style="font-family: Courier New; font-size: 20px;">PSCom.dll</span> on your system.

```vbscript
regasm %@truename[PSCom.dll] /codebase
```
In order to use PSCom from 32-bit VBScript,\
you need to create a dllsurrogate for <span style="font-family: Courier New; font-size: 20px;">PSCom.dll.</span>

I do this by running the <span style="font-family: Courier New; font-size: 20px;">surrogate.btm</span> file from Take Command Console;

```powershell
E:\...\Release>surrogate PSCom.PSScript 
{A5B64F74-0899-4475-853D-F81CB2024AE8}
Windows Registry Editor Version 5.00

[HKEY_CLASSES_ROOT\AppID\{A5B64F74-0899-4475-853D-F81CB2024AE8}]
"DllSurrogate"=""

[HKEY_CLASSES_ROOT\CLSID\{A5B64F74-0899-4475-853D-F81CB2024AE8}]
"AppID"="{A5B64F74-0899-4475-853D-F81CB2024AE8}"
```
I developed this COM ActiveX .dll using VB.NET in Visual Studio 2019 on Windows 10 Pro.

In the https://github.com/JoeC4281/PSCom/tree/master/bin/Release folder,\
there is the <span style="font-family: Courier New; font-size: 20px;">test.vbs</span> file;
```VB Script
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
```
Sample run;
```VB Script
E:\...\Release>cscript.exe //nologo test.vbs 
10.0.19045
19045
2025-04-02 2:11:43 PM
66
```
Here's a JScript example for using the PSCom.dll;
```VB Script
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
```


Here's a thinBasic example for using the PSCom.dll;
```VB Script
uses "Console"
' Uses "Trace"

dim ps as iDispatch
dim result as string

printl "Creating PSCom.PSScript object"

ps = CreateObject("PSCom.PSScript")

if IsComObject(ps) then
  result = ps.ExecuteScript("e:\utils\winver.ps1")

  printl result
  
  result = ps.ExecuteCommand("Get-Date; 2025-1959")
  
  Printl result
Else
  Printl "Could not create PSCom.PSScript object"
end if

ps = Nothing
```

Please consider this beta software.\
It may well have issues.\
Try it at your own risk.\
If you find a problem, please report it via **Issues** here in GitHub.

PSCom is currently licensed only for testing purposes.

This repository contains all of the source code so that you can compile it using Visual Studio 2019.

Take Command Console is from https://jpsoft.com/all-downloads/all-downloads.html


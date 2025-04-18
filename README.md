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

---
If you want to build the project from the command line,\
first make sure that you are using the <span style="font-family: Courier New; font-size: 20px;">**MSBuild.exe**</span> included with VS 2019.

On my Windows 10 system, I ran the <span style="font-family: Courier New; font-size: 20px;">**vcvarsall.bat**</span> from;

<span style="font-family: Courier New; font-size: 20px;">C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\VC\Auxiliary\Build\vcvarsall.bat</span>

This sets my development environment up properly for use with VS 2019 utils.

Next, I build the project;

<span style="font-family: Courier New; font-size: 20px;">**msbuild.exe PSCom.vbproj** **/p:Configuration=Release /p:Platform=AnyCPU**</span>

This creates <span style="font-family: Courier New; font-size: 20px;">**E:\Documents\vb.net\PSCom\bin\Release\PSCom.dll**</span>

---
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
```JavaScript
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
Here's a python example for using the PSCom.dll;
```Python
import win32com.client

# Create the COM object
obj = win32com.client.Dispatch("PSCom.PSScript")

# Define the Console class
class Console:
    def __init__(self):
        from win32com.client import Dispatch
        self.fso = Dispatch("Scripting.FileSystemObject")
        self.stdout = self.fso.GetStandardStream(1)  # Standard output
        self.stderr = self.fso.GetStandardStream(2)  # Standard error

    def Echo(self, theArg):
        self.stdout.WriteLine(theArg)

    def __del__(self):
        del self.stderr
        del self.stdout
        del self.fso

# Create an instance of the Console class
CScript = Console()

# Test ExecuteScript with a sample PowerShell script path
result = obj.ExecuteScript(r"E:\Utils\WinVer.ps1")
CScript.Echo(result)

# Test ExecuteCommand with a sample PowerShell command(s)
result = obj.ExecuteCommand("Get-Date; 2025-1959")
CScript.Echo(result)

# Clean up
del obj
```
Please consider this beta software.\
It may well have issues.\
Try it at your own risk.\
If you find a problem, please report it via **Issues** here in GitHub.

PSCom is currently licensed only for testing purposes.

This repository contains all of the source code so that you can compile it using Visual Studio 2019.

Take Command Console is from https://jpsoft.com/all-downloads/all-downloads.html


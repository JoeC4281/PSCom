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

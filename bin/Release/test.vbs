Dim obj
Dim result

' Create the COM object
Set obj = CreateObject("PSCom.PSScript")
Set CScript = New Console

' Test ExecuteScript with a sample PowerShell script path
result = obj.ExecuteScript("E:\Utils\WinVer.ps1")
CScript.Echo result

' Test ExecuteCommand with a sample PowerShell command(s)
result = obj.ExecuteCommand("Get-Date; 2025-1959")
CScript.Echo result

Set obj = Nothing

'set CScript = New Console
'CScript.Echo "This is VBScript"
Class Console
  Private fso
  Private stdout
  Private stderr

  Private Sub Class_Initialize()
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stdout = fso.GetStandardStream(1)
    Set stderr = fso.GetStandardStream(2)
  End Sub

  Public Sub Echo(theArg)
    stdout.WriteLine theArg
  End Sub

  Private Sub class_Terminate()
    Set stderr = Nothing
    Set stdout = Nothing
    Set fso = Nothing
  End Sub
End Class
Imports System.Management.Automation
Imports System.Runtime.InteropServices

<ComVisible(True)>
<Guid("A1285800-644C-4497-9478-DB42AF3634D7")>
Public Interface IPSScript
    Function ExecuteScript(scriptPath As String) As String
    Function ExecuteCommand(command As String) As String
End Interface

<ComVisible(True)>
<Guid("A5B64F74-0899-4475-853D-F81CB2024AE8")>
<ProgId("PSCom.PSScript")>
Public Class PSScript

    Implements IPSScript

    Public Function ExecuteScript(scriptPath As String) As String Implements IPSScript.ExecuteScript
        Try
            Using ps As PowerShell = PowerShell.Create()
                ' Read the PowerShell script from the file
                Dim scriptContent As String = System.IO.File.ReadAllText(scriptPath)
                ps.AddScript(scriptContent)
                Dim results = ps.Invoke()

                Dim output As String = String.Join(Environment.NewLine, results.Select(Function(r) r.ToString()))
                Return output
            End Using
        Catch ex As Exception
            Return $"Error: {ex.Message}"
        End Try
    End Function

    Public Function ExecuteCommand(command As String) As String Implements IPSScript.ExecuteCommand
        Try
            Using ps As PowerShell = PowerShell.Create()
                ' Add the script directly, not treating it as a file path
                ps.AddScript(command)
                Dim results = ps.Invoke()

                Dim output As String = String.Join(Environment.NewLine, results.Select(Function(r) r.ToString()))
                Return output
            End Using
        Catch ex As Exception
            Return $"Error: {ex.Message}"
        End Try
    End Function
End Class

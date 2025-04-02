Imports System.Management.Automation
Imports System.Runtime.InteropServices

<ComVisible(True)> ' Make class COM-visible
<Guid("A5B64F74-0899-4475-853D-F81CB2024AE8")> ' Unique GUID for COM class
<ProgId("PSCom.PSScript")>
Public Class PSScript
    Public Function ExecuteScript(scriptPath As String) As String
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

    Public Function Add2(theNumber As Long) As Long
        Return theNumber + 2
    End Function
End Class

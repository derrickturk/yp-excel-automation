Option Explicit
Implements ILogger

Public Sub ILogger_Log(ByVal msg As String)
    ActiveSheet.Range("A1").Value = msg
End Sub

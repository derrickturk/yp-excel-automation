Option Explicit
Implements ILogger

Public Sub ILogger_Log(ByVal msg As String)
    Debug.Print msg
End Sub

Option Explicit

Public Sub UseLogger()
    Dim logger As Ilogger
    Randomize
    If Rnd > 0.5 Then
        Set logger = New DebugLogger
    Else
        Set logger = New ExcelLogger
    End If
    logger.Log "test message"
End Sub

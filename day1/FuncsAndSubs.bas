Option Explicit

Public Sub WriteCell(ByVal cell As Range, ByVal contents As Variant)
    cell.Value = contents
End Sub

Public Function Factorial(ByVal x As Long) As Long
    If x <= 0 Then
        Factorial = 1
    Else
        Factorial = x * Factorial(x - 1)
    End If
End Function

Public Sub UseFunctions()
    Debug.Print "original value: " & ActiveSheet.Range("A1").Value
    WriteCell ActiveSheet.Range("A1"), 17
    Debug.Print "new contents: " & ActiveSheet.Range("A1").Value

    Debug.Print "factorial(5) = " & Factorial(5)
End Sub

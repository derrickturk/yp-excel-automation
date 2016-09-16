Option Explicit

Public Function Fn(ByVal x As Long, ByRef y As Long)
    x = x - 2
    y = x * 3
    Fn = y - x
End Function

Public Sub UseFn()
    Dim x As Long, y As Long, z As Long
    x = 5
    y = 7
    z = Fn(x, y)
    Debug.Print "x =", x, "y = ", y, "z = ", z
End Sub

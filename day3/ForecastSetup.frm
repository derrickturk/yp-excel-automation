Option Explicit

Private numericMonths_ As Long
Private okClicked_ As Boolean

Public Property Get ForecastMonths() As Long
    ForecastMonths = numericMonths_
End Property

Public Property Get OkClicked() As Boolean
    OkClicked = okClicked_
End Property

Private Sub UserForm_Initialize()
    numericMonths_ = 48
    okClicked_ = False
End Sub

Private Sub txtForecastMonths_Change()
    On Error GoTo INVALID
    numericMonths_ = CLng(txtForecastMonths.Value)
    Exit Sub
INVALID:
    txtForecastMonths.Value = ""
End Sub

Private Sub OkButton_Click()
    ForecastSetup.Hide
    okClicked_ = True
End Sub

Option Explicit

Implements IDecline

Private Const YearDays As Double = 365.25

Public qi As Double  ' initial daily rate [ units / day ]
Public Di As Double  ' initial decline [ nominal / year ]
Public b As Double   ' hyperbolic exponent [ unitless ]

' t is time in years
Public Function IDecline_Rate(ByVal t As Double)
    If b = 0 Then
        IDecline_Rate = qi * Exp(-Di * t)
    ElseIf b = 1 Then
        IDecline_Rate = qi / (1 + Di * t)
    Else
        IDecline_Rate = qi * (1 + b * Di * t) ^ (-1 / b)
    End If
End Function

' t is time in years
Public Function IDecline_Cumulative(ByVal t As Double)
    Dim yearlyRate As Double
    yearlyRate = qi * YearDays

    If Di = 0 Then
        IDecline_Cumulative = yearlyRate * t
    ElseIf b = 0 Then
        IDecline_Cumulative = yearlyRate / Di * (1 - Exp(-Di * t))
    ElseIf b = 1 Then
        IDecline_Cumulative = yearlyRate / Di * Log(1 + Di * t)
    Else
        IDecline_Cumulative = (yearlyRate / ((1 - b) * Di)) * _
          (1 - (1 + b * Di * t) ^ (1 - (1 / b)))
    End If
End Function

' tBegin, tEnd are time in years
Public Function IDecline_Volume(ByVal tBegin As Double, ByVal tEnd As Double) _ 
  As Double
    IDecline_Volume = IDecline_Cumulative(tEnd) - IDecline_Cumulative(tBegin)
End Function

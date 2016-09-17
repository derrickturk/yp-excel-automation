Option Explicit

Implements IDecline

Private Const YearDays As Double = 365.25

Public q1 As Double  ' initial daily rate [ units / day ]
Public a As Double   ' [ unitless ]
Public m As Double   ' [ unitless ]

' Duong t(a, b)
Private Function t_ab(ByVal t As Double)
    t_ab = (t ^ -m) * Exp(a * (t ^ (1 - m) - 1) / (1 - m))
End Function

' t is time in years
Public Function IDecline_Rate(ByVal t As Double)
    ' look, I didn't come up with the model, but it's terribly behaved
    '   at time = 0
    ' so here's a quick hack, because this isn't a webinar about the Duong
    '  model, it's an excuse to use polymorphism
    If t < 1 / YearDays Then
        IDecline_Rate = IDecline_Rate(1 / YearDays)
        Exit Function
    End If
    IDecline_Rate = q1 * t_ab(t)
End Function

' t is time in years
Public Function IDecline_Cumulative(ByVal t As Double)
    Dim yearlyRate As Double
    yearlyRate = IDecline_Rate(t) * YearDays

    ' see comment in IDecline_Rate above
    If t < 1 / YearDays Then
        IDecline_Cumulative = 0
        Exit Function
    End If

    IDecline_Cumulative = yearlyRate / (a * t ^ -m)
End Function

' tBegin, tEnd are time in years
Public Function IDecline_Volume(ByVal tBegin As Double, ByVal tEnd As Double) _ 
  As Double
    IDecline_Volume = IDecline_Cumulative(tEnd) - IDecline_Cumulative(tBegin)
End Function

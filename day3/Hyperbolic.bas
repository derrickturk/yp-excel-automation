Option Explicit

Public Type HyperbolicDecline
    qi As Double  ' initial daily rate [ units / day ]
    Di As Double  ' initial decline [ nominal / year ]
    b As Double   ' hyperbolic exponent [ unitless ]
End Type

' you'll notice this is slightly different than yesterday...
'   if this is defined prior to the HyperbolicDecline type, VBA begins
'   throwing a compile error when a variable of HyperbolicDecline type is
'   Dim'd in any other module.
' this naturally results in an extremely opaque "internal error" message
'   which, I kid you not, advises you to make sure you didn't accidentally
'   Err.Raise the error yourself, and then to call Microsoft support.
' and people wonder why I hate VBA so much...
Private Const YearDays As Double = 365.25

' t is time in years
Public Function Rate(ByRef decline As HyperbolicDecline, ByVal t As Double)
    If decline.b = 0 Then
        Rate = decline.qi * Exp(-decline.Di * t)
    ElseIf decline.b = 1 Then
        Rate = decline.qi / (1 + decline.Di * t)
    Else
        Rate = decline.qi * (1 + decline.b * decline.Di * t) ^ (-1 / decline.b)
    End If
End Function

' t is time in years
Public Function Cumulative(ByRef decline As HyperbolicDecline, _
  ByVal t As Double)
    Dim yearlyRate As Double
    yearlyRate = decline.qi * YearDays

    If decline.Di = 0 Then
        Cumulative = yearlyRate * t
    ElseIf decline.b = 0 Then
        Cumulative = yearlyRate / decline.Di * (1 - Exp(-decline.Di * t))
    ElseIf decline.b = 1 Then
        Cumulative = yearlyRate / decline.Di * Log(1 + decline.Di * t)
    Else
        Cumulative = (yearlyRate / ((1 - decline.b) * decline.Di)) * _
          (1 - (1 + decline.b * decline.Di * t) ^ (1 - (1 / decline.b)))
    End If
End Function

' tBegin, tEnd are time in years
Public Function Volume(ByRef decline As HyperbolicDecline, _
  ByVal tBegin As Double, ByVal tEnd As Double) As Double
    Volume = Cumulative(decline, tEnd) - Cumulative(decline, tBegin)
End Function

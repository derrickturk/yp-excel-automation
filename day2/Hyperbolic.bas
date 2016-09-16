Option Explicit
Option Base 0

' things to complain about here:
'   can't pass UDT ByRef (dumb. pass as const pointer. logical CBV-ness is
'     independent of ABI. and VB6 lets me pass objects ByVal despite the ABI
'     passing a pointer.)
'   ElseIf vs Else If

Private Const YearDays As Double = 365.25

Public Type HyperbolicDecline
    qi As Double  ' initial daily rate [ units / day ]
    Di As Double  ' initial decline [ nominal / year ]
    b As Double   ' hyperbolic exponent [ unitless ]
End Type

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

Public Sub TestHyperbolic
    Const Days As Long = 1000

    Dim decline As HyperbolicDecline
    decline.qi = 500
    decline.Di = 0.95
    decline.b = 0.75

    Dim forecastTime(0 To Days - 1) As Double
    Dim forecastRate(0 To Days - 1) As Double
    Dim forecastCumulative(0 To Days - 1) As Double

    Dim currentTime As Double
    currentTime = 0.0

    Dim i As Long
    For i = LBound(forecastTime) To UBound(forecastTime)
        forecastTime(i) = currentTime
        forecastRate(i) = Rate(decline, currentTime)
        forecastCumulative(i) = Cumulative(decline, currentTime)
        currentTime = currentTime + 1.0 / YearDays
    Next i

    ActiveSheet.Range( _
        ActiveSheet.Cells(1, 1), _
        ActiveSheet.Cells(Days, 1) _
    ) = Application.Transpose(forecastTime)

    ActiveSheet.Range( _
        ActiveSheet.Cells(1, 2), _
        ActiveSheet.Cells(Days, 2) _
    ) = Application.Transpose(forecastRate)

    ActiveSheet.Range( _
        ActiveSheet.Cells(1, 3), _
        ActiveSheet.Cells(Days, 3) _
    ) = Application.Transpose(forecastCumulative)
End Sub

Option Explicit

Public Function CreateDecline(ByVal declineType As String, _
  ByVal p1 As Double, ByVal p2 As Double, ByVal p3 As Double) As IDecline
    Select Case declineType
        Case "Hyperbolic"
            Dim hyp As HyperbolicDecline
            Set hyp = New HyperbolicDecline
            hyp.qi = p1
            hyp.Di = p2
            hyp.b = p3
            Set CreateDecline = hyp

        Case Else
            Set CreateDecline = Nothing
    End Select
End Function

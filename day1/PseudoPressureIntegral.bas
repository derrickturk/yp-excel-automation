Option Explicit

Private Const Viscosity As Double = 1.0
Private Const Z As Double = 0.997

Public Function PseudoPressure(ByVal p As Double, ByVal pBase As Double) _
  As Double
    PseudoPressure = 2 * TrapezoidalRule(pBase, p)
End Function

Private Function Integrand(ByVal p As Double)
    Integrand = p / (Viscosity * Z)
End Function

Private Function TrapezoidalRule(ByVal pLower as Double, _
  ByVal pUpper As Double) As Double
    Const Steps As Long = 100

    Dim integral As Double
    integral = 0

    Dim stepSize As Double
    stepSize = (pUpper - pLower) / Steps

    Dim pLeft As Double, pRight As Double
    pLeft = pLower
    pRight = pLower + stepSize

    Dim valLeft as Double, valRight as Double
    valLeft = Integrand(pLeft)
    valRight = Integrand(pRight)

    Dim i As Long
    For i = 1 To Steps
        integral = integral + 0.5 * stepSize * (valLeft + valRight)

        pLeft = pRight
        valLeft = valRight

        pRight = pLeft + stepSize
        valRight = Integrand(pRight)
    Next i

    TrapezoidalRule = integral
End Function

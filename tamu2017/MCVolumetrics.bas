Option Explicit

Private Const MaxAllowedPorosity As Double = 0.4

Public Type Realization
    Area As Double
    Height As Double
    Porosity As Double
    OilSaturation As Double
    FVF As Double
End Type

Public Type DistParams
    Area As Double
    MinHeight As Double
    MaxHeight As Double
    MeanPorosity As Double
    SDPorosity As Double
    MinOilSaturation As Double
    MaxOilSaturation As Double
    MinFVF As Double
    MaxFVF As Double
End Type

Public Function MCAverageOOIP( _
  ByVal Area As Double, _
  ByVal MinHeight As Double, _
  ByVal MaxHeight As Double, _
  ByVal MeanPorosity As Double, _
  ByVal SDPorosity As Double, _
  ByVal MinOilSaturation As Double, _
  ByVal MaxOilSaturation As Double, _
  ByVal MinFVF As Double, _
  ByVal MaxFVF As Double, _
  ByVal trials As Long, _
  ByVal seed As Long) _
  As Double
    Dim params As DistParams
    params.Area = Area
    params.MinHeight = MinHeight
    params.MaxHeight = MaxHeight
    params.MeanPorosity = MeanPorosity
    params.SDPorosity = SDPorosity
    params.MinOilSaturation = MinOilSaturation
    params.MaxOilSaturation = MaxOilSaturation
    params.MinFVF = MinFVF
    params.MaxFVF = MaxFVF

    MCAverageOOIP = Average(OOIPForSamples(MonteCarlo(trials, params, seed)))
End Function

Public Function Average(ByRef values() As Double) As Double
    Dim runningAvg As Long
    runningAvg = 0

    Dim scalar As Double
    scalar = 1# / (UBound(values) - LBound(values) + 1)

    Dim i As Long
    For i = LBound(values) To UBound(values)
        runningAvg = values(i) * scalar + runningAvg
    Next i

    Average = runningAvg
End Function

Public Function OOIPForSamples(ByRef samples() As Realization) _
  As Double()
    Dim result() As Double
    ReDim result(LBound(samples) To UBound(samples))
    Dim i As Long
    For i = LBound(samples) To UBound(samples)
        result(i) = OOIP(samples(i))
    Next i
    OOIPForSamples = result
End Function

Public Function OOIP(ByRef sample As Realization) As Double
    OOIP = 7758 * _
           sample.Area * _
           sample.Height * _
           sample.Porosity * _
           sample.OilSaturation _
       / sample.FVF
End Function

Public Function MonteCarlo(ByVal trials As Long, _
  ByRef params As DistParams, ByVal seed As Long) As Realization()
    Dim result() As Realization
    ReDim result(1 To trials) As Realization

    Rnd (-1#)
    Randomize seed

    Dim i As Long
    For i = 1 To trials
        result(i) = MonteCarloRealization(params)
    Next i
    MonteCarlo = result
End Function

Private Function MonteCarloRealization(ByRef params As DistParams) _
  As Realization
    Dim result As Realization

    result.Area = params.Area

    result.Height = BoundedRandUniform( _
      params.MinHeight, params.MaxHeight)

    result.Porosity = RandNormal( _
      params.MeanPorosity, params.SDPorosity)
    If result.Porosity < 0 Then
        result.Porosity = 0
    ElseIf result.Porosity > MaxAllowedPorosity Then
        result.Porosity = MaxAllowedPorosity
    End If

    result.OilSaturation = BoundedRandUniform( _
      params.MinOilSaturation, params.MaxOilSaturation)

    result.FVF = BoundedRandUniform( _
      params.MinFVF, params.MaxFVF)

    MonteCarloRealization = result
End Function

' transform a uniformly distributed value on [0, 1) to
'   a normally distributed value using the Box-Muller transform;
'   we're not going to use the second value we could generate
'   for simplicity's sake
Private Function RandNormal(ByVal mean As Double, _
  ByVal sd As Double) As Double
    Dim u1 As Double
    Dim u2 As Double
    u1 = Rnd()
    u2 = Rnd()
    Dim r As Double
    r = Sqr(-2# * Log(u1))
    Dim theta As Double
    theta = 2# * Application.WorksheetFunction.Pi * u2

    ' n.b. we forgot part of this in the workshop;
    '   if we don't scale by sd and shift by mean,
    '   we just have a draw from the standard normal distribution
    '   ~ N (0, 1)
    RandNormal = mean + sd * r * Cos(theta)
End Function

Private Function BoundedRandUniform(ByVal lower As Double, _
  ByVal upper As Double) As Double
    BoundedRandUniform = Rnd() * (upper - lower) + lower
End Function

Public Sub TestMC()
    Dim params As DistParams
    params.Area = 40
    params.MinHeight = 10
    params.MaxHeight = 50
    params.MeanPorosity = 0.1
    params.SDPorosity = 0.01
    params.MinOilSaturation = 0.3
    params.MaxOilSaturation = 0.5
    params.MinFVF = 1.2
    params.MaxFVF = 1.5

    Dim samples() As Realization
    samples = MonteCarlo(10000, params, 12345)

    Dim ooips() As Double
    ooips = OOIPForSamples(samples)

    Dim avg As Double
    avg = Average(ooips)

    Debug.Print avg
End Sub

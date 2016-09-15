Option Explicit
Option Base 0

Private Type DataRecord
    DateTime As Date
    Pressure As Double
    Temperature As Double
End Type

Private Function RandomRecords(ByVal days As Long) As DataRecord()
    Const TempRange As Double = 5
    Const PressureRange As Double = 10

    Dim records() As DataRecord
    ReDim records(0 To (days * 24) - 1)

    Dim CurrentDate As Date
    CurrentDate = #1/1/2016 12:00 AM#

    Dim baseTemp As Double, basePressure As Double
    baseTemp = 70
    basePressure = 100

    Randomize

    Dim i As Long
    For i = LBound(records) To UBound(records)
        records(i).DateTime = CurrentDate
        records(i).Pressure = UniformRandom(basePressure - PressureRange, _
          basePressure + PressureRange)
        records(i).Temperature = UniformRandom(baseTemp - TempRange, _
          baseTemp + TempRange)

        CurrentDate = DateAdd("h", 1, CurrentDate)
        baseTemp = baseTemp + 2 / 24
        basePressure = basePressure - 5 / 24
    Next i

    RandomRecords = records
End Function

Private Function UniformRandom(ByVal lb As Double, ByVal ub As Double) As Double
    UniformRandom = lb + Rnd * (ub - lb)
End Function

Public Sub WriteRandomData(ByVal days As Long, ByVal sheet)
    Dim records() As DataRecord
    records = RandomRecords(days)

    ' note: "array-like" objects returned from Excel are almost always 1-based
    '   so Cells(1, 1) is the first row, first column cell 
    '   and Cells(0, 0) does not exist
    sheet.Cells(1, 1).Value = "Date/Time"
    sheet.Cells(1, 2).Value = "Pressure"
    sheet.Cells(1, 3).Value = "Temperature"

    Dim row As Long
    row = 2

    Dim i As Long
    For i = LBound(records) To UBound(records)
        sheet.Cells(row, 1).Value = records(i).DateTime
        sheet.Cells(row, 2).Value = records(i).Pressure
        sheet.Cells(row, 3).Value = records(i).Temperature
        row = row + 1
    Next i
End Sub

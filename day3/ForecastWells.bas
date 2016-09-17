Option Explicit

Private Const DataBeginRow As Long = 2
Private Const WellNameColumn As Long = 1
Private Const BeginParamsColumn As Long = 2

Private Const ForecastMonths As Long = 48
Private Const YearMonths = 12

Public Sub test()
    ForecastDeclines Sheet1
End Sub

Public Sub ForecastDeclines(ByVal sheet As Worksheet)
    Dim lastCell As Range
    Set lastCell = LastUsedCell(sheet)

    Dim prodWorkbook As Workbook
    Set prodWorkbook = Workbooks.Add()
    Dim initialSheets As Long
    initialSheets = prodWorkbook.Worksheets.Count

    Dim row As Long
    For row = DataBeginRow To lastCell.Row
        Dim wellName As String
        wellName = sheet.Cells(row, WellNameColumn).Value

        Dim wellDecline As HyperbolicDecline
        wellDecline.qi = sheet.Cells(row, BeginParamsColumn).Value
        wellDecline.Di = sheet.Cells(row, BeginParamsColumn + 1).Value
        wellDecline.b = sheet.Cells(row, BeginParamsColumn + 2).Value

        Dim volumes(0 To ForecastMonths - 1) As Double
        Dim elapsedTime As Double
        elapsedTime = 0
        Dim i As Long
        For i = LBound(volumes) To UBound(volumes)
            volumes(i) = Volume(wellDecline, _
              elapsedTime, elapsedTime + 1 / YearMonths)
            elapsedTime = elapsedTime + 1 / YearMonths
        Next i

        Dim wellSheet As Worksheet
        Set wellSheet = prodWorkbook.Sheets.Add( _
          After := prodWorkbook.Sheets(prodWorkbook.Sheets.Count))

        FormatProductionSheet wellSheet, wellName, volumes
    Next row

    If row < DataBeginRow Then ' we didn't have any records
        prodWorkbook.Close False ' don't save changes
        MsgBox "No records!"
        Exit Sub
    End If

    ' we don't want the default Excel message box to prompt for
    '   confirmation when we delete these (empty) sheets
    Application.DisplayAlerts = False
    For i = 1 To initialSheets
        prodWorkbook.Sheets(1).Delete
    Next i
    Application.DisplayAlerts = True

    Dim prodWorkbookFilename As String
    prodWorkbookFilename = ProductionWorkbookFilename()
    If prodWorkbookFilename = "" Then ' user cancelled "Save As" dialog
        prodWorkbook.Close False
        Exit Sub
    End If

    prodWorkbook.SaveAs prodWorkbookFilename
    prodWorkbook.Close
End Sub

Private Sub FormatProductionSheet(ByVal sheet As Worksheet, _
  ByVal wellName As String, ByRef volumes() As Double)
    sheet.Name = wellName

    sheet.Range("A1").Value = wellName
    sheet.Range("A1").Font.Bold = True
    sheet.Range("A1:B1").Merge
    sheet.Range("A2").Value = "Month"
    sheet.Range("A2").Font.Bold = True
    sheet.Range("B2").Value = "Volume"
    sheet.Range("B2").Font.Bold = True

    sheet.Range( _
        sheet.Cells(3, 1), _
        sheet.Cells(3 + ForecastMonths - 1, 1) _
    ).Value = Application.Transpose(Sequence(1, ForecastMonths))

    sheet.Range( _
        sheet.Cells(3, 2), _
        sheet.Cells(3 + ForecastMonths - 1, 2) _
    ).Value = Application.Transpose(volumes)
End Sub

Private Function Sequence(ByVal seqFrom As Long, ByVal seqTo As Long) As Long()
    Dim seq() As Long
    ReDim seq(seqFrom to seqTo)
    Dim i As Long
    For i = LBound(seq) To UBound(seq)
        seq(i) = i
    Next i
    Sequence = seq
End Function

Private Function ProductionWorkbookFilename() As String
    Dim result As Variant
    result = Application.GetSaveAsFilename( _
        InitialFilename := "monthly_production.xlsx", _
        FileFilter := "Excel workbooks (*.xlsx), *.xlsx", _
        Title := "Save Monthly Production Workbook")
    If result <> False Then ' sic
        ProductionWorkbookFilename = result
    Else
        ProductionWorkbookFilename = ""
    End If
End Function

' returns last occupied cell of worksheet (or .Cells(1, 1) if empty)
Private Function LastUsedCell(ByVal sheet As Worksheet) As Range
    Dim lastRow As Long, lastCol As Long
    If Application.WorksheetFunction.CountA(sheet.Cells) <> 0 Then
        lastRow = sheet.Cells.Find(What := "*", After := sheet.Range("A1"), _
            LookAt := xlPart, LookIn := xlFormulas, _
            SearchOrder := xlByRows, SearchDirection := xlPrevious, _
            MatchCase := False).Row
        lastCol = sheet.Cells.Find(What := "*", After := sheet.Range("A1"), _
            LookAt := xlPart, LookIn := xlFormulas, _
            SearchOrder := xlByColumns, SearchDirection := xlPrevious, _
            MatchCase := False).Column
    End If

    ' fun VBA fact: if you leave off "set" here; you'll get a runtime
    '   error when you try to return the default property of the Range object,
    '   which is its .Value; this is of course to ensure a user-friendly
    '   and welcoming experience for beginners
    Set LastUsedCell = sheet.Cells(lastRow, lastCol)
End Function

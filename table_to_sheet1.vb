Sub CreateFilteredTablesForEachCenter()
    Dim mainSheet As Worksheet, newSheet As Worksheet
    Dim electionCenters As Object, centerName As Variant
    Dim i As Long, lastRow As Long, tbl As ListObject
    Dim rngToFilter As Range, filteredRange As Range

    On Error GoTo ErrorHandler

    ' Initialize settings
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set mainSheet = ActiveSheet
    Set electionCenters = CreateObject("Scripting.Dictionary")

    ' Define data range (adjust if your data is larger)
    lastRow = mainSheet.Cells(mainSheet.Rows.Count, "A").End(xlUp).Row
    Set rngToFilter = mainSheet.Range("A1:B" & lastRow)

    ' Collect unique election centers
    For i = 2 To lastRow
        If Not IsEmpty(mainSheet.Cells(i, 2)) Then
            centerName = mainSheet.Cells(i, 2).Value
            If Not electionCenters.Exists(centerName) And centerName <> "" Then
                electionCenters.Add centerName, 1
            End If
        End If
    Next i

    ' Exit if no centers found
    If electionCenters.Count = 0 Then
        MsgBox "No election centers found.", vbExclamation
        GoTo CleanUp
    End If

    ' Process each center
    For Each centerName In electionCenters.Keys
        ' Create new sheet
        Set newSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        newSheet.DisplayRightToLeft = True

        ' Set sheet name
        On Error Resume Next
        newSheet.Name = Left(centerName, 31)
        If Err.Number <> 0 Then newSheet.Name = "Center_" & electionCenters.Count
        On Error GoTo 0

        ' Copy filtered data
        rngToFilter.AutoFilter Field:=2, Criteria1:=centerName
        Set filteredRange = rngToFilter.SpecialCells(xlCellTypeVisible)
        filteredRange.Copy newSheet.Range("A1")
        rngToFilter.AutoFilter

        ' Convert to formatted table
        Set tbl = newSheet.ListObjects.Add(xlSrcRange, newSheet.Range("A1").CurrentRegion, , xlYes)
        tbl.Name = "Table_" & Replace(centerName, " ", "_")
        tbl.TableStyle = "TableStyleMedium2"

        ' Add summary row with formulas
        tbl.ShowTotals = True
        tbl.ListColumns(1).TotalsCalculation = xlTotalsCalculationCount
        tbl.ListColumns(2).TotalsCalculation = xlTotalsCalculationNone

        ' Format the table
        With newSheet
            .Cells.HorizontalAlignment = xlRight
            .Columns("A:B").AutoFit
            ' .Cells(1, 3).Value = "Report Generated: " & Now()
            .Cells(1, 3).Font.Italic = True
        End With
    Next centerName

CleanUp:
    mainSheet.Activate
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    If electionCenters.Count > 0 Then
        MsgBox "Created " & electionCenters.Count & " professional center reports", vbInformation, "Success"
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
    Resume CleanUp
End Sub
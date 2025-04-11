Sub CreateSheetsForEachElectionCenter_RTL()
    Dim mainSheet As Worksheet
    Dim newSheet As Worksheet
    Dim electionCenters As Object
    Dim i As Long
    Dim centerName As Variant
    Dim filteredRange As Range
    Dim rngToFilter As Range

    On Error GoTo ErrorHandler

    ' Turn off screen updating for better performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Set the main sheet (assuming the table is on the active sheet)
    Set mainSheet = ActiveSheet

    ' Set the range to filter (A1:B5)
    Set rngToFilter = mainSheet.Range("A1:B5")

    ' Create dictionary to store unique election centers
    Set electionCenters = CreateObject("Scripting.Dictionary")

    ' Collect all unique election centers from column B (B2 to B5)
    For i = 2 To 5 ' Assuming row 1 has headers, data is in rows 2-5
        If Not IsEmpty(mainSheet.Cells(i, 2)) Then
            centerName = mainSheet.Cells(i, 2).Value ' Column B contains centers
            If Not electionCenters.Exists(centerName) And centerName <> "" Then
                electionCenters.Add centerName, 1
            End If
        End If
    Next i

    ' Exit if no centers found
    If electionCenters.Count = 0 Then
        MsgBox "No election centers found in the specified range.", vbExclamation
        GoTo CleanUp
    End If

    ' Delete old sheets if they exist (except the main sheet)
    For Each newSheet In ThisWorkbook.Worksheets
        If newSheet.Name <> mainSheet.Name And electionCenters.Exists(newSheet.Name) Then
            Application.DisplayAlerts = False
            newSheet.Delete
            Application.DisplayAlerts = True
        End If
    Next newSheet

    ' Create a new sheet for each election center
    For Each centerName In electionCenters.Keys
        ' Create new sheet
        Set newSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))

        ' Set sheet to right-to-left
        newSheet.DisplayRightToLeft = True

        On Error Resume Next
        newSheet.Name = Left(centerName, 31) ' Ensure sheet name doesn't exceed 31 chars
        If Err.Number <> 0 Then
            newSheet.Name = "Center_" & electionCenters.Count
            Err.Clear
        End If
        On Error GoTo 0

        ' Copy headers from main sheet
        mainSheet.Range("A1:B1").Copy Destination:=newSheet.Range("A1")

        ' Filter and copy data for this center
        With rngToFilter
            .AutoFilter
            .AutoFilter Field:=2, Criteria1:=centerName

            ' Check if any visible cells after filtering
            On Error Resume Next
            Set filteredRange = .Offset(1, 0).Resize(.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
            On Error GoTo 0

            If Not filteredRange Is Nothing Then
                filteredRange.Copy Destination:=newSheet.Range("A2")
                Set filteredRange = Nothing
            End If

            .AutoFilter
        End With

        ' Auto-fit columns to content
        newSheet.Columns("A:B").AutoFit

        ' Add voter count in cell C1 (right-aligned)
        ' اضافة عدد الناخبين الى الخلية
        ' With newSheet.Cells(1, 3)
        '     .Value = "عدد الناخبين: " & (newSheet.Cells(newSheet.Rows.Count, "A").End(xlUp).Row - 1)
        '     .HorizontalAlignment = xlRight
        ' End With

        ' Set entire sheet to right-to-left
        newSheet.Cells.HorizontalAlignment = xlRight
    Next centerName

CleanUp:
    ' Return to main sheet
    mainSheet.Activate

    ' Restore screen updating
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    If electionCenters.Count > 0 Then
        MsgBox "Done  workly by Eng. sajjad " & electionCenters.Count & " this code work good !", vbInformation, "done"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "check data and try again .", vbCritical, "error big problem"
    Resume CleanUp
End Sub
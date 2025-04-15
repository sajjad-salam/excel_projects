Sub CountVoters()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim totalPeople As Long, countTahdeeth As Long, countTamTahdeeth As Long
    Dim resultSheet As Worksheet, updateSheet As Worksheet
    Dim outputRow As Long, updateRow As Long
    Dim cellValue As String
    Dim responsibleName As String

    ' Create or clear results sheet
    Application.ScreenUpdating = False

    On Error Resume Next
    Set resultSheet = ThisWorkbook.Sheets("Results")
    Set updateSheet = ThisWorkbook.Sheets("غير المحدثين")
    On Error GoTo 0

    ' Prepare Results sheet
    If resultSheet Is Nothing Then
        Set resultSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        resultSheet.Name = "Results"
    Else
        resultSheet.Cells.Clear
    End If

    ' Prepare غير المحدثين sheet
    If updateSheet Is Nothing Then
        Set updateSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        updateSheet.Name = "غير المحدثين"
    Else
        updateSheet.Cells.Clear
    End If

    ' Set up headers in Arabic
    resultSheet.Range("A1:D1").Value = Array("اسم المسؤول", "عدد التحديث", "عدد تم التحديث", "عدد البطايق")
    outputRow = 2

    ' Set up headers for غير المحدثين sheet
    updateSheet.Range("A1:C1").Value = Array("ت", "اسم الناخب الثلاثي", "المسؤول عنه")
    updateSheet.Range("A1:C1").Font.Bold = True
    updateSheet.Range("A1:C1").HorizontalAlignment = xlCenter
    updateSheet.Range("A1:C1").Interior.Color = RGB(191, 191, 191)
    updateRow = 2

    ' Loop through all sheets
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Results" And ws.Name <> "غير المحدثين" Then
            countTahdeeth = 0
            countTamTahdeeth = 0
            totalPeople = 0

            ' Remove "10_ورقة1" from sheet name for responsible person
            responsibleName = Replace(ws.Name, "10_ورقة1", "")

            ' Find last row in column C (المركز الانتخابي)
            lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
            If lastRow < 9 Then lastRow = 9 ' Ensure at least row 9 is checked
            If lastRow > 100 Then lastRow = 100 ' Don't go beyond row 100

            ' Loop through data rows (9 to lastRow)
            For i = 9 To lastRow
                If Not IsEmpty(ws.Cells(i, 3).Value) Then ' Check column C
                    totalPeople = totalPeople + 1
                    cellValue = Trim(ws.Cells(i, 3).Value)

                    ' Check for تم التحديث (exact match first)
                    If cellValue = "تم التحديث" Then
                        countTamTahdeeth = countTamTahdeeth + 1
                    ' Check for تحديث (exact match)
                    ElseIf cellValue = "تحديث" Then
                        countTahdeeth = countTahdeeth + 1
                        ' Add to غير المحدثين sheet
                        updateSheet.Cells(updateRow, 1).Value = updateRow - 1 ' Sequence number
                        updateSheet.Cells(updateRow, 2).Value = ws.Cells(i, 2).Value ' Voter name
                        updateSheet.Cells(updateRow, 3).Value = responsibleName ' Responsible sheet (cleaned)
                        updateRow = updateRow + 1
                    ' Check for partial matches if needed
                    ElseIf InStr(1, cellValue, "تم التحديث", vbTextCompare) > 0 Then
                        countTamTahdeeth = countTamTahdeeth + 1
                    ElseIf InStr(1, cellValue, "تحديث", vbTextCompare) > 0 Then
                        countTahdeeth = countTahdeeth + 1
                        ' Add to غير المحدثين sheet
                        updateSheet.Cells(updateRow, 1).Value = updateRow - 1 ' Sequence number
                        updateSheet.Cells(updateRow, 2).Value = ws.Cells(i, 2).Value ' Voter name
                        updateSheet.Cells(updateRow, 3).Value = responsibleName ' Responsible sheet (cleaned)
                        updateRow = updateRow + 1
                    End If
                End If
            Next i

            ' Write results to Results sheet
            resultSheet.Cells(outputRow, 1).Value = ws.Name
            resultSheet.Cells(outputRow, 2).Value = countTahdeeth
            resultSheet.Cells(outputRow, 3).Value = countTamTahdeeth
            resultSheet.Cells(outputRow, 4).Value = totalPeople - countTahdeeth - countTamTahdeeth

            outputRow = outputRow + 1
        End If
    Next ws

    ' Format Results sheet
    With resultSheet
        .Columns("A:D").AutoFit
        .Range("A1:D1").Font.Bold = True
        .Range("A1:D1").HorizontalAlignment = xlCenter
        .Range("A1:D1").Interior.Color = RGB(191, 191, 191)

        ' Add borders
        If outputRow > 2 Then
            .Range("A1:D" & outputRow - 1).Borders.LineStyle = xlContinuous
        End If

        ' Add totals row
        If outputRow > 2 Then
            .Cells(outputRow, 1).Value = "المجموع"
            .Cells(outputRow, 2).Formula = "=SUM(B2:B" & outputRow - 1 & ")"
            .Cells(outputRow, 3).Formula = "=SUM(C2:C" & outputRow - 1 & ")"
            .Cells(outputRow, 4).Formula = "=SUM(D2:D" & outputRow - 1 & ")"
            .Rows(outputRow).Font.Bold = True
        End If
    End With

    ' Format غير المحدثين sheet
    With updateSheet
        If updateRow > 2 Then
            .Range("A1:C" & updateRow - 1).Borders.LineStyle = xlContinuous
            .Columns("A:C").AutoFit
            .Columns("B").ColumnWidth = 30 ' Wider column for names
            .Columns("C").ColumnWidth = 25 ' Wider column for responsible names
        Else
            .Range("A2:C2").Value = Array("", "لا يوجد ناخبين غير محدثين", "")
        End If
    End With

    Application.ScreenUpdating = True
    MsgBox "تم الانتهاء من عملية العد بنجاح!" & vbNewLine & _
           "النتائج موجودة في ورقة 'Results'" & vbNewLine & _
           "قائمة غير المحدثين موجودة في ورقة 'غير المحدثين'", vbInformation, "اكتمل"
End Sub

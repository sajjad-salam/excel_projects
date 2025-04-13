Sub MergeExcelFilesToSheets()
    Dim folderPath As String
    Dim outputFileName As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim wbSource As Workbook
    Dim wbDest As Workbook
    Dim ws As Worksheet
    Dim sheetName As String
    Dim fileCount As Integer

    ' Initialize variables
    fileCount = 0
    outputFileName = "Merged_Workbook.xlsx"

    ' Get folder path from user
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder Containing Excel Files"
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            MsgBox "No folder selected. Operation cancelled.", vbExclamation
            Exit Sub
        End If
    End With

    ' Check if folder exists
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        MsgBox "The folder does not exist!", vbExclamation
        Exit Sub
    End If

    ' Create a new workbook for output
    Set wbDest = Workbooks.Add
    Application.DisplayAlerts = False

    ' Loop through all Excel files in the folder
    Set folder = fso.GetFolder(folderPath)
    For Each file In folder.Files
        ' Check if file is Excel file (xls, xlsx, xlsm) and not temporary
        If (Right(file.Name, 4) = ".xls" Or Right(file.Name, 5) = ".xlsx" Or Right(file.Name, 5) = ".xlsm") _
           And Left(file.Name, 2) <> "~$" Then

            ' Open source workbook
            Set wbSource = Workbooks.Open(file.Path, ReadOnly:=True)

            ' Copy each sheet from source to destination workbook
            For Each ws In wbSource.Worksheets
                ' Create sheet name (max 31 chars, no special characters)
                sheetName = Left(file.Name, InStrRev(file.Name, ".") - 1)
                ' sheetName = Left(file.Name, InStrRev(file.Name, ".") - 1) & "_" & ws.Name
                sheetName = Left(sheetName, 31)
                sheetName = Replace(sheetName, ":", "")
                sheetName = Replace(sheetName, "\", "")
                sheetName = Replace(sheetName, "/", "")
                sheetName = Replace(sheetName, "?", "")
                sheetName = Replace(sheetName, "*", "")
                sheetName = Replace(sheetName, "[", "")
                sheetName = Replace(sheetName, "]", "")

                ' Copy sheet to destination workbook
                ws.Copy After:=wbDest.Sheets(wbDest.Sheets.Count)
                wbDest.Sheets(wbDest.Sheets.Count).Name = sheetName
            Next ws

            ' Close source workbook without saving
            wbSource.Close SaveChanges:=False
            fileCount = fileCount + 1
        End If
    Next file

    ' Remove the default Sheet1 if it's empty
    On Error Resume Next
    If wbDest.Sheets("Sheet1").UsedRange.Address = "$A$1" And _
       wbDest.Sheets("Sheet1").Range("A1").Value = "" Then
        Application.DisplayAlerts = False
        wbDest.Sheets("Sheet1").Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0

    ' Save the merged workbook
    wbDest.SaveAs folderPath & "\" & outputFileName, FileFormat:=xlOpenXMLWorkbook
    wbDest.Close

    ' Show completion message
    MsgBox "Successfully merged " & fileCount & " files into " & outputFileName & ".", vbInformation

    ' Clean up
    Set fso = Nothing
    Set folder = Nothing
    Set file = Nothing
    Set wbSource = Nothing
    Set wbDest = Nothing
End Sub

Sub FilterAndCreateSheetsByColumnG()
    Dim wsMain As Worksheet
    Dim wsNew As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim filterRange As Range
    Dim uniqueValues As Collection
    Dim cell As Range
    Dim value As Variant
    Dim newSheetName As String

    ' تعطيل تحديث الشاشة لتحسين الأداء
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' تحديد الورقة الرئيسية (الشيت الذي يحتوي على البيانات)
    Set wsMain = ActiveSheet

    ' تحديد آخر صف وآخر عمود في الجدول
    lastRow = wsMain.Cells(wsMain.Rows.Count, "G").End(xlUp).Row
    lastCol = wsMain.Cells(1, wsMain.Columns.Count).End(xlToLeft).Column

    ' تحديد النطاق الكامل للبيانات (بما في ذلك العناوين)
    Set filterRange = wsMain.Range(wsMain.Cells(1, 1), wsMain.Cells(lastRow, lastCol))

    ' جمع القيم الفريدة من العمود G
    Set uniqueValues = New Collection
    On Error Resume Next
    For Each cell In wsMain.Range("G1:G" & lastRow) ' افتراض أن الصف الأول عناوين
        If Not IsEmpty(cell.Value) Then
            uniqueValues.Add cell.Value, CStr(cell.Value) ' إضافة القيمة إذا كانت فريدة
        End If
    Next cell
    On Error GoTo 0

    ' التحقق مما إذا كانت هناك قيم فريدة
    If uniqueValues.Count = 0 Then
        MsgBox "لم يتم العثور على أي قيم فريدة في العمود G!", vbExclamation
        GoTo CleanUp
    End If

    ' حذف الأوراق القديمة التي تحمل نفس الأسماء (اختياري)
    Dim sheet As Worksheet
    For Each sheet In ThisWorkbook.Worksheets
        If sheet.Name <> wsMain.Name Then
            Application.DisplayAlerts = False
            sheet.Delete
            Application.DisplayAlerts = True
        End If
    Next sheet

    ' إنشاء أوراق عمل جديدة لكل قيمة فريدة
    For Each value In uniqueValues
        ' إنشاء ورقة جديدة
        Set wsNew = Worksheets.Add(After:=Worksheets(Worksheets.Count))

        ' تعيين اسم الورقة (تجنب الأسماء الطويلة أو غير الصالحة)
        newSheetName = Left(value, 31)
        wsNew.Name = newSheetName

        ' نسخ العناوين إلى الورقة الجديدة
        filterRange.Rows(1).Copy Destination:=wsNew.Range("A1")

        ' تصفية البيانات بناءً على القيمة الحالية
        filterRange.AutoFilter Field:=7, Criteria1:=value ' العمود السابع (G)
        filterRange.Offset(1, 0).Resize(filterRange.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Copy _
            Destination:=wsNew.Range("A2")

        ' إزالة التصفية
        filterRange.AutoFilter
    Next value

CleanUp:
    ' إعادة تفعيل تحديث الشاشة والإشعارات
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ' إشعار بالانتهاء
    MsgBox "تم إنشاء " & uniqueValues.Count & " أوراق جديدة بنجاح!", vbInformation, "عملية ناجحة"
End Sub
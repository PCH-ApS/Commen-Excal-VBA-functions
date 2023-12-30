Attribute VB_Name = "CommonFunctions"
'@IgnoreModule DefaultMemberRequired
'@Folder "Code Budget"
Option Explicit













Public Function RemoveColumnsWithEmptyHeaderFromArray(ByRef arr As Variant) As Variant

    Dim arrColumns As Long
    Dim columnName As String
    Dim newColumnCount As Long
    Dim keepColumnsCount As Long
    Dim keepColumns() As Variant
    
    keepColumnsCount = 1
    For arrColumns = LBound(arr, 2) To UBound(arr, 2)
        columnName = arr(1, arrColumns)
        If Not columnName = vbNullString Then
            newColumnCount = newColumnCount + 1
            ReDim Preserve keepColumns(1 To keepColumnsCount)
            keepColumns(keepColumnsCount) = columnName
            keepColumnsCount = keepColumnsCount + 1
        End If
    Next arrColumns
    
    keepColumnsCount = 0
    
    Dim tempArr() As Variant
    ReDim tempArr(LBound(arr, 1) To UBound(arr, 1), 1 To newColumnCount) As Variant

    Dim arrRows As Long
    For arrColumns = LBound(arr, 2) To UBound(arr, 2)
        For keepColumnsCount = LBound(keepColumns) To UBound(keepColumns)
            If StrConv(CStr(arr(1, arrColumns)), vbUpperCase) = StrConv(CStr(keepColumns(keepColumnsCount)), vbUpperCase) Then
                For arrRows = LBound(arr, 1) To UBound(arr, 1)
                    tempArr(arrRows, keepColumnsCount) = arr(arrRows, arrColumns)
                Next arrRows
            End If
        Next keepColumnsCount
    Next arrColumns
  
    RemoveColumnsWithEmptyHeaderFromArray = tempArr

End Function



Public Function CreateArrayFromTable(ByRef tbl As ListObject) As Variant

    Dim numRows As Long
    Dim numCols As Long
    numRows = tbl.range.rows.Count
    numCols = tbl.range.columns.Count

    Dim myArray() As Variant
    ReDim myArray(1 To numRows, 1 To numCols)

    Dim i As Long, j As Long
    For i = 1 To numRows
        For j = 1 To numCols
            myArray(i, j) = tbl.range.Cells(i, j).Value
        Next j
    Next i
    
    CreateArrayFromTable = myArray
    
End Function

Public Function CreateArrayFromCSV(ByVal reportFile As String, delimitter As String) As Variant

    Open reportFile For Input As #1
    Dim numCols As Long
    Dim numRows As Long
    Dim rowFromFile As String
    Dim columnFromRow() As String
        
    Do Until EOF(1)
        Line Input #1, rowFromFile
        columnFromRow = Split(rowFromFile, delimitter)
        numRows = numRows + 1
        'Array starts from 0 thus + 1. Check all rows to find max number of columns
        If UBound(columnFromRow) + 1 > numCols Then numCols = numCols + 1
    Loop
    
    Dim myArray() As Variant
    ReDim myArray(1 To numRows, 1 To numCols)
    Close #1
    
    Open reportFile For Input As #1
    numRows = 0
    Dim j As Long
    Do Until EOF(1)
        Line Input #1, rowFromFile
        columnFromRow = Split(rowFromFile, delimitter)
        numRows = numRows + 1
        For j = 0 To UBound(columnFromRow)
            myArray(numRows, j + 1) = columnFromRow(j)
        Next j
    Loop
    Close #1
    
    CreateArrayFromCSV = myArray
    
End Function




Public Function SelectSingleFile(ByVal dialogTitle As String, ByVal fileFilter As String, ByVal fileType As String) As String
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = dialogTitle
        .Filters.Clear
        .Filters.Add fileFilter, fileType
        If .Show = True Then
                SelectSingleFile = .SelectedItems(1)
            Else
                Exit Function
        End If
    End With

End Function
Public Function SelectMultiFile(ByVal dialogTitle As String, ByVal fileFilter As String, ByVal fileType As String) As Variant

    Dim fncSelectedItem() As Variant
    Dim filesSelected As Long
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = True
        .Title = dialogTitle
        .Filters.Clear
        .Filters.Add fileFilter, fileType
        If .Show = True Then
                ReDim fncSelectedItem(1 To .SelectedItems.Count)
                For filesSelected = 1 To .SelectedItems.Count
                    fncSelectedItem(filesSelected) = .SelectedItems(filesSelected)
                Next
                SelectMultiFile = fncSelectedItem()
            Else
                Exit Function
        End If
    End With

End Function
Public Function SelectSaveDest(ByVal filename As String, ByVal fileFilter As String, ByVal dialogTitle As String) As String
    
    SelectSaveDest = Application.GetSaveAsFilename(fileFilter:=fileFilter, Title:=dialogTitle, InitialFileName:=filename)
    
End Function
Public Sub ModuleStart(Optional ByVal ActWB As Workbook)

    On Error Resume Next
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    If ActWB Is Nothing Then
        ThisWorkbook.AutoSaveOn = False
    Else
        ActWB.AutoSaveOn = False
    End If
    On Error GoTo 0
    
End Sub
Public Sub ModuleEnd(Optional ByVal ActWB As Workbook)

    On Error Resume Next
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    If ActWB Is Nothing Then
        ThisWorkbook.AutoSaveOn = True
    Else
        ActWB.AutoSaveOn = True
    End If
    On Error GoTo 0
    
End Sub





Public Function CheckListObjectExists(ByVal ListObjTCheck As String, ByVal WksToCheck As Worksheet) As Boolean
    
    Dim NewListObjTCheck As String
    If InStr(ListObjTCheck, " ") > 0 Then NewListObjTCheck = Replace(ListObjTCheck, " ", "_")
    If NewListObjTCheck = vbNullString Then NewListObjTCheck = ListObjTCheck
    Dim fncLo As ListObject
    On Error Resume Next
    Set fncLo = WksToCheck.ListObjects(StrConv(NewListObjTCheck, vbUpperCase))
    On Error GoTo 0
    CheckListObjectExists = Not fncLo Is Nothing
    
End Function
Public Sub RemoveEmptyLine(ByVal wksSub As Worksheet)
    
    Dim wksListObj As ListObject
    Set wksListObj = wksSub.ListObjects(wksSub.Name)
    wksListObj.DataBodyRange.AutoFilter field:=wksListObj.ListColumns(1).Index, Criteria1:="="
    If wksListObj.ListColumns(1).range.SpecialCells(xlVisible).Count > 1 Then
        Application.DisplayAlerts = False
        wksListObj.DataBodyRange.rows.Delete
        Application.DisplayAlerts = True
    End If
    wksListObj.AutoFilter.ShowAllData
    
End Sub
Public Sub DisplayProgressSingle(pctCompl1 As Long, pctText1 As String, timeText3 As String)

CommonSingleProgress.Text.Caption = pctText1 & ": " & Int(pctCompl1) & "%"
CommonSingleProgress.Bar.Width = pctCompl1 * 2
CommonSingleProgress.Text3.Caption = "Elapsed time: " & timeText3
DoEvents

End Sub
Public Sub DisplayProgressOverall(pctCompl1 As Long, pctCompl2 As Long, pctText1 As String, pctText2 As String, timeText3 As String)

CommonOverallProgress.Text.Caption = pctText1 & ": " & Int(pctCompl1) & "%"
CommonOverallProgress.Bar.Width = pctCompl1 * 2
CommonOverallProgress.Text2.Caption = pctText2 & ": " & Int(pctCompl2) & "%"
CommonOverallProgress.Bar2.Width = pctCompl2 * 2
CommonOverallProgress.Text3.Caption = "Elapsed time: " & timeText3
DoEvents

End Sub
Public Function ConvertRangeToArray(ByVal range As range) As Variant

    Dim resultArray() As String
    Dim resultRows As Long
    Dim resultColumns As Long
    
    resultRows = range.rows.Count
    resultColumns = range.columns.Count
    
    ReDim resultArray(resultRows, resultColumns)
    resultArray = range
    
    ConvertRangeToArray = resultArray
    
End Function
Public Function CheckColumnExistsInArray(ByVal ColumnToCheck As String, ByRef ArrayToCheck() As Variant) As Boolean
    
    Dim i As Long
    Dim colunmFoundInPos As Long
    
    For i = LBound(ArrayToCheck, 2) To UBound(ArrayToCheck, 2)
        colunmFoundInPos = 0
        If UCase(ArrayToCheck(1, i)) = UCase(ColumnToCheck) Then
            colunmFoundInPos = i
            Exit For
        End If
    Next
    
    If colunmFoundInPos > 0 Then
        CheckColumnExistsInArray = True
    Else
        CheckColumnExistsInArray = False
    End If
    
End Function
Public Function UniqueValuesFromArray(ByRef ArrayToCheck() As Variant, ByVal columnName As String) As Variant

    Dim i As Long
    Dim arrayColumn As Long
    Dim arrayRow As Long
    Dim sheetsList() As String
    Dim sheetsListUnique As New Collection, a
    
    i = i + 1
    Do While Not ArrayToCheck(1, i) = vbNullString
        If ArrayToCheck(1, i) = columnName Then
            arrayColumn = i
            Exit Do
        End If
        i = i + 1
    Loop

    On Error Resume Next
    For i = 2 To UBound(ArrayToCheck) ' **** avoid headr row #1
        If Not ArrayToCheck(i, arrayColumn) = vbNullString Then
            a = CStr(ArrayToCheck(i, arrayColumn))
            sheetsListUnique.Add a, a
        End If
    Next
    On Error GoTo 0
    ReDim sheetsList(1 To sheetsListUnique.Count)
    For i = 1 To sheetsListUnique.Count
        sheetsList(i) = sheetsListUnique(i)
    Next i
    UniqueValuesFromArray = sheetsList

End Function
Public Function UniqueValuesFromTable(ByVal WksToCheck As Worksheet, ByVal ListObjToCheck As String, columnName As String) As Variant

    Dim Source_tbl As ListObject
    Dim sourceColumn As ListColumn
    Dim sheetsList() As String
    Dim sheetsListUnique As New Collection, a
    Dim i As Long
      
    Set Source_tbl = WksToCheck.ListObjects(ListObjToCheck)
    Set sourceColumn = Source_tbl.ListColumns(columnName)
    If Not sourceColumn Is Nothing Then
        On Error Resume Next
        For Each a In sourceColumn.DataBodyRange
            If Not a = vbNullString Then sheetsListUnique.Add a, a
        Next
        On Error GoTo 0
        ReDim sheetsList(1 To sheetsListUnique.Count)
        For i = 1 To sheetsListUnique.Count
            sheetsList(i) = sheetsListUnique(i)
        Next i
        UniqueValuesFromTable = sheetsList
    End If

End Function
Function GetFolder(ByVal dialogTitle As String) As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = dialogTitle
        .AllowMultiSelect = False
        '.InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function
Public Sub RemoveBlankTableRows(ByVal wksSub As Worksheet, ByVal tblName As String)
    
    Dim tbl As ListObject
    Dim tblRow As ListRow
    Dim rowStr As String
    Dim rows As Long
    Dim columns As Long
    Dim i As Long
    Dim j As Long
    Dim datarows As Long
    Dim emptyrows As Long
    
    Set tbl = wksSub.ListObjects(tblName)
       
    datarows = 1
    emptyrows = 0
    rows = tbl.ListRows.Count
    columns = tbl.ListColumns.Count
    For i = rows To 1 Step -1
        rowStr = vbNullString
        For j = 1 To columns
            rowStr = rowStr & CStr(tbl.ListColumns(j).DataBodyRange(i)) & ";"
        Next j
        If Not Len(rowStr) > columns Then
            Set tblRow = tbl.ListRows(i)
            tblRow.Delete
        End If
    Next i
    
End Sub
Public Sub CloseOpenWorkbook(filename As String)

    Dim closeWb As Workbook
    
    On Error Resume Next
    If Not filename = vbNullString Then
        Set closeWb = Workbooks.Open(filename, UpdateLinks:=0, Local:=True)
        If Not closeWb Is Nothing Then closeWb.Close (False)
    End If
    On Error GoTo 0

End Sub
Public Function RemoveColumnsFromArray(ByRef arr As Variant, colToKeep As Variant) As Variant
      
    Dim tempArr() As Variant
    ReDim tempArr(LBound(arr, 1) To UBound(arr, 1), LBound(colToKeep) To UBound(colToKeep)) As Variant
    
    Dim arrRows As Long
    Dim keepColumns As Long
    Dim arrColumns As Long
    
    For arrColumns = LBound(arr, 2) To UBound(arr, 2)
        For keepColumns = LBound(colToKeep) To UBound(colToKeep)
            If StrConv(CStr(arr(1, arrColumns)), vbUpperCase) = StrConv(CStr(colToKeep(keepColumns)), vbUpperCase) Then
                For arrRows = LBound(arr, 1) To UBound(arr, 1)
                    tempArr(arrRows, keepColumns) = arr(arrRows, arrColumns)
                Next arrRows
            End If
        Next keepColumns
    Next arrColumns
  
    RemoveColumnsFromArray = tempArr
    
End Function


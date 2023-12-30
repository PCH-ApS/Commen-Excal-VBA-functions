Attribute VB_Name = "CommonFunctions"
'@IgnoreModule DefaultMemberRequired
'@Folder "Code Budget"
Option Explicit
Public Sub SetDateInTableColumn(ByRef sourceTable As ListObject, columnTitle As String)

    Dim sourceColumn As ListColumn
    Set sourceColumn = sourceTable.ListColumns(columnTitle)
    Dim a As Object
    If Not sourceColumn Is Nothing Then
        For Each a In sourceColumn.DataBodyRange
            If IsError(a) Then a = Now
        Next
    End If
    
End Sub
Public Function SourceSearchStringFoundInTargetArrayV2(ByRef sourceSearchString As String, ByRef targetArray As Variant, ByRef headerArray As Variant, ByVal searchKey As Variant) As Boolean
  
    Dim targetRowCount As Long
    Dim targetCompareString As String
    Dim targetColumnCount As Long
    Dim headerRowCount As Long
    
    For targetRowCount = LBound(targetArray, 1) + 1 To UBound(targetArray, 1)
        If targetArray(targetRowCount, 1) = searchKey Then
            targetCompareString = vbNullString
            For targetColumnCount = LBound(targetArray, 2) To UBound(targetArray, 2)
                For headerRowCount = LBound(headerArray, 1) To UBound(headerArray, 1)
                    If Not InStr(headerArray(headerRowCount, 1), "*") > 0 Then
                        If headerArray(headerRowCount, 1) = targetArray(1, targetColumnCount) Then
                            targetCompareString = targetCompareString & ";" & CStr(targetArray(targetRowCount, targetColumnCount))
                            Exit For
                        End If
                    End If
                Next headerRowCount
            Next targetColumnCount
            If targetCompareString = sourceSearchString Then
                SourceSearchStringFoundInTargetArrayV2 = True
                Exit Function
            End If
        End If
    Next targetRowCount
    
    SourceSearchStringFoundInTargetArrayV2 = False

End Function
Public Function CompareSourceWithTargetArrayAndCountRowsToAddFromSourceArrayV2(ByRef sourceArray As Variant, ByRef targetArray As Variant, ByRef headerArray As Variant, ByVal dictionaryColumn As String) As Variant

' **** Get pos of dictionary column in target array
    Dim headerRowCount As Long
    Dim dictionaryColumnNo As Long
    For headerRowCount = LBound(headerArray, 1) To UBound(headerArray, 1)
        If dictionaryColumn = headerArray(headerRowCount, 1) Then
            dictionaryColumnNo = headerArray(headerRowCount, 2)
        End If
    Next headerRowCount

' **** Compare arrays and add rows from importArray
    Dim newRowCount As Long
    newRowCount = 1
    Dim idArray() As Long
    ReDim Preserve idArray(1 To newRowCount)

' **** Create Dictionary of dates from targetArray
    Dim targetDates As Object
    Set targetDates = CreateObject("Scripting.Dictionary")
    Dim targetRowCount As Long
    For targetRowCount = LBound(targetArray, 1) + 1 To UBound(targetArray, 1)
        If Not targetDates.Exists(targetArray(targetRowCount, dictionaryColumnNo)) Then
            targetDates.Add targetArray(targetRowCount, dictionaryColumnNo), 1
        End If
    Next targetRowCount

' **** Get pos of dictionary column in source array
    For headerRowCount = LBound(headerArray, 1) To UBound(headerArray, 1)
        If dictionaryColumn = headerArray(headerRowCount, 1) Then
            dictionaryColumnNo = headerArray(headerRowCount, 3)
        End If
    Next headerRowCount

' **** Loop though sourceArray, skip header row
    Dim importRowCount As Long
    Dim sourceCompareString As String
    Dim importColumnCount As Long
    For importRowCount = LBound(sourceArray, 1) + 1 To UBound(sourceArray, 1)
    ' **** Test if the date is found in targetArray, if not found, then assume row should be added
        If targetDates.Exists(sourceArray(importRowCount, dictionaryColumnNo)) Then
            sourceCompareString = vbNullString
            For importColumnCount = LBound(sourceArray, 2) To UBound(sourceArray, 2)
                For headerRowCount = LBound(headerArray, 1) To UBound(headerArray, 1)
                    If Not InStr(headerArray(headerRowCount, 1), "*") > 0 Then
                        If headerArray(headerRowCount, 1) = sourceArray(1, importColumnCount) Then
                            sourceCompareString = sourceCompareString & ";" & CStr(sourceArray(importRowCount, importColumnCount))
                            Exit For
                        End If
                    End If
                Next headerRowCount
            Next importColumnCount
            If CommonFunctions.SourceSearchStringFoundInTargetArrayV2(sourceCompareString, targetArray, headerArray, sourceArray(importRowCount, 1)) = False Then
                ReDim Preserve idArray(1 To newRowCount)
                idArray(newRowCount) = importRowCount
                newRowCount = newRowCount + 1
            End If
        Else
            ReDim Preserve idArray(1 To newRowCount)
            idArray(newRowCount) = importRowCount
            newRowCount = newRowCount + 1
        End If
    Next importRowCount
    
    CompareSourceWithTargetArrayAndCountRowsToAddFromSourceArrayV2 = idArray

End Function

Public Function GetArrayColumnHeadersFromSourceAndTargetArray(ByRef sourceArray As Variant, ByRef targetArray As Variant) As Variant

    Dim headerArray() As Variant
    ReDim headerArray(1 To UBound(targetArray, 2), 1 To 3)
    Dim targetRowCount As Long
    For targetRowCount = LBound(targetArray, 2) To UBound(targetArray, 2)
        headerArray(targetRowCount, 1) = targetArray(1, targetRowCount)
        headerArray(targetRowCount, 2) = targetRowCount
    Next targetRowCount
    
    Dim importRowCount As Long
    Dim headerArrayRowCount As Long
    For importRowCount = LBound(sourceArray, 2) To UBound(sourceArray, 2)
        For headerArrayRowCount = LBound(headerArray, 1) To UBound(headerArray, 1)
            If sourceArray(1, importRowCount) = headerArray(headerArrayRowCount, 1) Then
                headerArray(headerArrayRowCount, 3) = importRowCount
                Exit For
            End If
        Next headerArrayRowCount
    Next importRowCount
    
    GetArrayColumnHeadersFromSourceAndTargetArray = headerArray

End Function
Public Function SourceSearchStringFoundInTargetArray(ByRef sourceSearchString As String, ByRef targetArray As Variant, headerArray As Variant) As Boolean
  
    Dim targetRowCount As Long
    Dim targetColumnCount As Long
    Dim headerRowCount As Long
    Dim targetCompareString As String

    For targetRowCount = LBound(targetArray, 1) + 1 To UBound(targetArray, 1)
        targetCompareString = vbNullString
        For targetColumnCount = LBound(targetArray, 2) To UBound(targetArray, 2)
            For headerRowCount = LBound(headerArray, 1) To UBound(headerArray, 1)
                If Not InStr(headerArray(headerRowCount, 1), "*") > 0 Then
                    If headerArray(headerRowCount, 1) = targetArray(1, targetColumnCount) Then
                        targetCompareString = targetCompareString & ";" & CStr(targetArray(targetRowCount, targetColumnCount))
                        Exit For
                    End If
                End If
            Next headerRowCount
        Next targetColumnCount
        If targetCompareString = sourceSearchString Then
            SourceSearchStringFoundInTargetArray = True
            Exit Function
        End If
    Next targetRowCount

    SourceSearchStringFoundInTargetArray = False

End Function
Public Function AddMissingRowArrayToTargetArray(ByRef sourceArray As Variant, ByRef targetArray As Variant, headerArray As Variant, sourceRowIDarr As Variant) As Variant
    
    Dim newRowCount As Long
    newRowCount = 1
    Dim idArray() As Long
    Dim targetRowCount As Long
    For targetRowCount = LBound(targetArray, 1) To UBound(targetArray, 1)
        If targetArray(targetRowCount, 1) = vbNullString Then
            ReDim Preserve idArray(1 To newRowCount)
            idArray(newRowCount) = targetRowCount
            newRowCount = newRowCount + 1
        End If
    Next targetRowCount

    Dim targetRowID As Long
    If UBound(idArray) = UBound(sourceRowIDarr) Then
        Dim sourceRowID As Long
        Dim headerRowCount As Long
        For targetRowCount = LBound(idArray) To UBound(idArray)
            targetRowID = idArray(targetRowCount)
            sourceRowID = sourceRowIDarr(targetRowCount)
            For headerRowCount = LBound(headerArray, 1) To UBound(headerArray, 1)
                If Not InStr(headerArray(headerRowCount, 1), "*") > 0 Then
                    targetArray(targetRowID, headerArray(headerRowCount, 2)) = sourceArray(sourceRowID, headerArray(headerRowCount, 3))
                End If
            Next headerRowCount
        Next targetRowCount
    End If

    AddMissingRowArrayToTargetArray = targetArray

End Function
Public Function CompareSourceWithTargetArrayAndCountRowsToAddFromSourceArray(ByRef sourceArray As Variant, ByRef targetArray As Variant, ByRef headerArray As Variant) As Variant

' **** Compare arrays and add rows from importArray
    Dim importRowCount As Long
    Dim sourceCompareString As String
    Dim importColumnCount As Long
    Dim headerRowCount As Long
    Dim newRowCount As Long
    newRowCount = 1
    Dim idArray() As Long
    ReDim Preserve idArray(1 To newRowCount)

' **** Count new rows to add to
    For importRowCount = LBound(sourceArray, 1) + 1 To UBound(sourceArray, 1)
        sourceCompareString = vbNullString
        For importColumnCount = LBound(sourceArray, 2) To UBound(sourceArray, 2)
            For headerRowCount = LBound(headerArray, 1) To UBound(headerArray, 1)
                If Not InStr(headerArray(headerRowCount, 1), "*") > 0 Then
                    If headerArray(headerRowCount, 1) = sourceArray(1, importColumnCount) Then
                        sourceCompareString = sourceCompareString & ";" & CStr(sourceArray(importRowCount, importColumnCount))
                        Exit For
                    End If
                End If
            Next headerRowCount
        Next importColumnCount
        If CommonFunctions.SourceSearchStringFoundInTargetArray(sourceCompareString, targetArray, headerArray) = False Then
            ReDim Preserve idArray(1 To newRowCount)
            idArray(newRowCount) = importRowCount
            newRowCount = newRowCount + 1
        End If
    Next importRowCount
    
    CompareSourceWithTargetArrayAndCountRowsToAddFromSourceArray = idArray

End Function

Public Function CreateLargerArrayFromExistingArray(ByRef arr As Variant, ByVal startPos As Long, ByVal extraRows As Long, Optional ByVal extraColumns As Long) As Variant

    Dim counter As Long
    Dim existingRowsCount As Long
    Dim existingColumnCount As Long
    
    For counter = LBound(arr, 1) To UBound(arr, 1)
        existingRowsCount = existingRowsCount + 1
    Next counter
    
    For counter = LBound(arr, 2) To UBound(arr, 2)
        existingColumnCount = existingColumnCount + 1
    Next counter
    
    Dim myArray() As Variant
    If Not extraColumns = 0 Then
        ReDim myArray(startPos To (existingRowsCount + extraRows), startPos To (existingColumnCount + extraColumns))
    Else
        ReDim myArray(startPos To (existingRowsCount + extraRows), startPos To existingColumnCount)
    End If
    
    Dim counterColumn As Long
    For counter = LBound(arr, 1) To UBound(arr, 1)
        For counterColumn = LBound(arr, 2) To UBound(arr, 2)
            myArray(counter, counterColumn) = arr(counter, counterColumn)
        Next counterColumn
    Next counter
    
    CreateLargerArrayFromExistingArray = myArray
    
End Function
Public Function ConvertTextToDateInArray(ByRef arr As Variant, ByVal dateColumnName As String, oldDateSeperator As String, newDateSeperator As String) As Variant
   
    Dim columnCount As Long
    Dim columnName As String
    Dim rowCount As Long
    Dim cellValue As Variant
    Dim tmpCellValue As Variant
    
    For columnCount = LBound(arr, 2) To UBound(arr, 2)
        columnName = arr(1, columnCount)
        If columnName = dateColumnName Then
            For rowCount = LBound(arr, 1) + 1 To UBound(arr, 1)
                cellValue = arr(rowCount, columnCount)
                If Not IsDate(cellValue) Then
                    tmpCellValue = Replace(cellValue, oldDateSeperator, newDateSeperator)
                    If IsDate(tmpCellValue) Then
                        arr(rowCount, columnCount) = tmpCellValue
                        arr(rowCount, columnCount) = CDate(arr(rowCount, columnCount))
                    End If
                Else
                    tmpCellValue = CDate(cellValue)
                    If IsDate(tmpCellValue) Then
                        tmpCellValue = Replace(cellValue, oldDateSeperator, newDateSeperator)
                        arr(rowCount, columnCount) = CDate(cellValue)
                    End If
                End If
            Next rowCount
        End If
    Next columnCount
    
    ConvertTextToDateInArray = arr
    
End Function
Public Function ConvertTextToNumberInArray(ByRef arr As Variant, Optional decimalSeperator As String) As Variant
   
    Dim columnCount As Long
    Dim rowCount As Long
    Dim cellValue As Variant
    Dim tmpCellValue As Variant
    For columnCount = LBound(arr, 2) To UBound(arr, 2)
        For rowCount = LBound(arr, 1) To UBound(arr, 1)
            cellValue = arr(rowCount, columnCount)
            If IsNumeric(cellValue) Then
                If Not decimalSeperator = vbNullString Then
                    If InStr(cellValue, decimalSeperator) > 0 And (Len(cellValue) - InStr(cellValue, decimalSeperator) <= 2) Then
                        tmpCellValue = Replace(cellValue, decimalSeperator, ",")
                    End If
                Else
                    tmpCellValue = cellValue
                End If
                tmpCellValue = CDec(tmpCellValue)
                arr(rowCount, columnCount) = tmpCellValue
            End If
        Next rowCount
    Next columnCount
    
    ConvertTextToNumberInArray = arr

End Function
Public Function RemoveDoubleQuotesFromArray(ByRef arr As Variant) As Variant
   
    Dim columnCount As Long
    Dim rowCount As Long
    Dim cellString As String
    
    For columnCount = LBound(arr, 2) To UBound(arr, 2)
        For rowCount = LBound(arr, 1) To UBound(arr, 1)
            cellString = arr(rowCount, columnCount)
            If InStr(1, cellString, """", vbTextCompare) > 0 Then
                Dim tmpCellString As String
                '@Ignore EmptyStringLiteral
                tmpCellString = Replace(cellString, """", "")
                arr(rowCount, columnCount) = Trim(tmpCellString)
            End If
        Next rowCount
    Next columnCount
    
    RemoveDoubleQuotesFromArray = arr

End Function
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

Public Function RemoveEmptyRowsFromArray(ByRef arr As Variant) As Variant
    
    Dim i As Long
    Dim j As Long
    Dim tempIndex As Long
    
    'Redim preserve only works on the LAST dimention, so need to know how many data lines there are
    For i = LBound(arr, 1) To UBound(arr, 1)
    
        Dim allColumnsEmpty As Boolean
        allColumnsEmpty = True
        For j = LBound(arr, 2) To UBound(arr, 2)
            If Len(arr(i, j)) > 0 Then
                allColumnsEmpty = False
                Exit For
            End If
        Next j
        
        If Not allColumnsEmpty Then tempIndex = tempIndex + 1
        
    Next i
    
    'If any blank lines was found tempIndex will be smaller then the original array
    If tempIndex < UBound(arr, 1) Then
    
       Dim tempArr() As Variant
       ReDim tempArr(LBound(arr, 1) To tempIndex, LBound(arr, 2) To UBound(arr, 2))
       
       tempIndex = 0
       
       For i = LBound(arr, 1) To UBound(arr, 1)
           allColumnsEmpty = True
           For j = LBound(arr, 2) To UBound(arr, 2)
               If Len(arr(i, j)) > 0 Then
                   allColumnsEmpty = False
                   Exit For
               End If
           Next j
           
           If Not allColumnsEmpty Then
               tempIndex = tempIndex + 1
               
               For j = LBound(arr, 2) To UBound(arr, 2)
                   tempArr(tempIndex, j) = arr(i, j)
               Next j
           End If
            
       Next i
    End If
    
    If tempIndex = 0 Then
        'all rows were empty, so set array to empty
        Erase arr
    End If
    
    If tempIndex < UBound(arr, 1) Then
        RemoveEmptyRowsFromArray = tempArr
        Exit Function
    End If
  
    RemoveEmptyRowsFromArray = arr
    
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
Public Function CreateArrayFromString(ByVal valueString As String) As Variant

'Count string elements
    Dim columnFromRow() As String
    columnFromRow = Split(valueString, ";")
    Dim i As Double
    Dim elements As Double
    For i = 0 To UBound(columnFromRow)
         elements = elements + 1
    Next i
    ReDim myArray(1 To elements)
    For i = LBound(columnFromRow) To UBound(columnFromRow)
        myArray(i + 1) = columnFromRow(i)
    Next i
    
    CreateArrayFromString = myArray
     
End Function

Public Function LogOutput(ByVal fncWksCodeName As String, ByVal fncLstObj As String, ByVal fncNew As Boolean) As Variant

    Dim fncLogOutputwks As Worksheet
    Dim fncLogOutputTbl As ListObject
    Dim columns As Long
    
    For Each fncLogOutputwks In ThisWorkbook.Worksheets
        If fncLogOutputwks.CodeName = fncWksCodeName Then
            Set fncLogOutputTbl = fncLogOutputwks.ListObjects(fncLstObj)
            Exit For
        End If
    Next

    On Error Resume Next
    fncLogOutputTbl.AutoFilter.ShowAllData
    On Error GoTo 0

    If fncNew = True Then
        If fncLogOutputTbl.DataBodyRange.rows.Count > 1 Then
            fncLogOutputTbl.DataBodyRange.Offset(1, 0).Resize(fncLogOutputTbl.DataBodyRange.rows.Count - 1, fncLogOutputTbl.DataBodyRange.columns.Count).rows.Delete
            For columns = 1 To fncLogOutputTbl.ListColumns("Tekst").Index
                fncLogOutputTbl.ListColumns(columns).DataBodyRange.Clear
            Next columns
        End If
    End If
    Set LogOutput = fncLogOutputTbl
    
End Function
Public Sub LogOutputLines(ByVal subTbl As ListObject, Optional ByVal subCtrlTxt As String, Optional ByVal subLogKey As String, Optional ByVal subLogTxt As String)
    
    Dim TblRows As Long
    Dim TblDate As Date
    Dim TblColumn As Long
    
    TblRows = subTbl.DataBodyRange.rows.Count
    TblDate = Now
    If Not subTbl.ListColumns("#").DataBodyRange(TblRows) = vbNullString Then
        subTbl.ListRows.Add
        TblRows = subTbl.DataBodyRange.rows.Count
    End If

    On Error Resume Next
    subTbl.AutoFilter.ShowAllData
    On Error GoTo 0
    
    For TblColumn = 1 To subTbl.ListColumns.Count
        If "#" = subTbl.ListColumns(TblColumn).Name Then subTbl.ListColumns("#").DataBodyRange(TblRows) = TblRows
        If "Dato" = subTbl.ListColumns(TblColumn).Name Then subTbl.ListColumns("Dato").DataBodyRange(TblRows) = CStr(TblDate)
        If "Handling" = subTbl.ListColumns(TblColumn).Name Then subTbl.ListColumns("Handling").DataBodyRange(TblRows) = subCtrlTxt
        If "Nøgle" = subTbl.ListColumns(TblColumn).Name Then subTbl.ListColumns("Nøgle").DataBodyRange(TblRows) = subLogKey
        If "Tekst" = subTbl.ListColumns(TblColumn).Name Then subTbl.ListColumns("Tekst").DataBodyRange(TblRows) = subLogTxt
    Next TblColumn
    
End Sub
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
Public Sub RemoveInitialSheet(Optional ByVal subWb As Workbook)
    
    Application.DisplayAlerts = False
    Dim subWks As Worksheet
    If subWb Is Nothing Then
        For Each subWks In ThisWorkbook.Worksheets
            If subWks.CodeName = "Sheet1" Then
                subWks.Delete
            End If
        Next
    Else
        For Each subWks In subWb.Worksheets
            If subWks.CodeName = "Sheet1" Then
                subWks.Delete
            End If
        Next
    End If

    Application.DisplayAlerts = True
    
End Sub
Public Sub CreateTable(ByVal wksSub As Worksheet)
    
    Dim rngSub As range
    Dim tablename As String

    tablename = StrConv(wksSub.Name, vbUpperCase)
    If InStr(tablename, " ") > 0 Then tablename = Replace(tablename, " ", "_")

    Set rngSub = wksSub.UsedRange
    If wksSub.ListObjects.Count < 1 Then
        wksSub.ListObjects.Add(xlSrcRange, rngSub, , xlYes).Name = tablename
        wksSub.ListObjects(tablename).TableStyle = vbNullString
    End If
    
End Sub
Public Sub RemoveTable(ByVal WbWithListObj As Workbook)
    
    Dim SheetWithListObj As Worksheet
    Dim listobj As ListObject
    Dim ListToRemove As Boolean
    
    For Each SheetWithListObj In WbWithListObj.Worksheets
        With SheetWithListObj
            ListToRemove = False
            For Each listobj In .ListObjects
                If StrConv(listobj.Name, vbUpperCase) = StrConv(.Name, vbUpperCase) Then ListToRemove = True
                If ListToRemove = True Then SheetWithListObj.ListObjects(StrConv(.Name, vbUpperCase)).Unlist
            Next listobj
        End With
    Next SheetWithListObj
    
End Sub
Public Function CheckColumnExists(ByVal ColumnToCheck As String, ByVal ListObjToCheck As ListObject) As Boolean

    Dim ColumnToTest As ListColumn
    On Error Resume Next
    Set ColumnToTest = ListObjToCheck.ListColumns(ColumnToCheck)
    On Error GoTo 0
    CheckColumnExists = Not ColumnToTest Is Nothing
    
End Function
Public Function CheckSheetExists(ByVal NameToCheck As String, Optional ByVal WbToCheck As Workbook) As Boolean

    If WbToCheck Is Nothing Then
        Set WbToCheck = ThisWorkbook
    End If

    Dim WksToCheck As Worksheet
    On Error Resume Next
    Set WksToCheck = WbToCheck.Sheets(NameToCheck)
    On Error GoTo 0
    CheckSheetExists = Not WksToCheck Is Nothing
    
End Function
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


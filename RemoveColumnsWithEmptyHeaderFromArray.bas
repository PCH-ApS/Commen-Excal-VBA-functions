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
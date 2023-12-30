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
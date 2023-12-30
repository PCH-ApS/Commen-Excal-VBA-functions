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
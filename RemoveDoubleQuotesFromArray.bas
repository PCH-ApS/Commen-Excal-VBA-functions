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
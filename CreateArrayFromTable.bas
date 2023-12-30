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
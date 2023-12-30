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
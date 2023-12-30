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
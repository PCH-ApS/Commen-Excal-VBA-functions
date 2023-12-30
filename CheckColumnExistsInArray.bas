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
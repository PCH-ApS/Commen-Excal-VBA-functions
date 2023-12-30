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
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
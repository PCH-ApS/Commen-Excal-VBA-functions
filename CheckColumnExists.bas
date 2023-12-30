Public Function CheckColumnExists(ByVal ColumnToCheck As String, ByVal ListObjToCheck As ListObject) As Boolean

    Dim ColumnToTest As ListColumn
    On Error Resume Next
    Set ColumnToTest = ListObjToCheck.ListColumns(ColumnToCheck)
    On Error GoTo 0
    CheckColumnExists = Not ColumnToTest Is Nothing
    
End Function
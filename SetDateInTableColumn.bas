Public Sub SetDateInTableColumn(ByRef sourceTable As ListObject, columnTitle As String)

    Dim sourceColumn As ListColumn
    Set sourceColumn = sourceTable.ListColumns(columnTitle)
    Dim a As Object
    If Not sourceColumn Is Nothing Then
        For Each a In sourceColumn.DataBodyRange
            If IsError(a) Then a = Now
        Next
    End If
    
End Sub
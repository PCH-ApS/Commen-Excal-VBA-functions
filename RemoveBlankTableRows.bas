Public Sub RemoveBlankTableRows(ByVal wksSub As Worksheet, ByVal tblName As String)
    
    Dim tbl As ListObject
    Dim tblRow As ListRow
    Dim rowStr As String
    Dim rows As Long
    Dim columns As Long
    Dim i As Long
    Dim j As Long
    Dim datarows As Long
    Dim emptyrows As Long
    
    Set tbl = wksSub.ListObjects(tblName)
       
    datarows = 1
    emptyrows = 0
    rows = tbl.ListRows.Count
    columns = tbl.ListColumns.Count
    For i = rows To 1 Step -1
        rowStr = vbNullString
        For j = 1 To columns
            rowStr = rowStr & CStr(tbl.ListColumns(j).DataBodyRange(i)) & ";"
        Next j
        If Not Len(rowStr) > columns Then
            Set tblRow = tbl.ListRows(i)
            tblRow.Delete
        End If
    Next i
    
End Sub
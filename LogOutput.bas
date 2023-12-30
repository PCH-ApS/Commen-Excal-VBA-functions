Public Function LogOutput(ByVal fncWksCodeName As String, ByVal fncLstObj As String, ByVal fncNew As Boolean) As Variant

    Dim fncLogOutputwks As Worksheet
    Dim fncLogOutputTbl As ListObject
    Dim columns As Long
    
    For Each fncLogOutputwks In ThisWorkbook.Worksheets
        If fncLogOutputwks.CodeName = fncWksCodeName Then
            Set fncLogOutputTbl = fncLogOutputwks.ListObjects(fncLstObj)
            Exit For
        End If
    Next

    On Error Resume Next
    fncLogOutputTbl.AutoFilter.ShowAllData
    On Error GoTo 0

    If fncNew = True Then
        If fncLogOutputTbl.DataBodyRange.rows.Count > 1 Then
            fncLogOutputTbl.DataBodyRange.Offset(1, 0).Resize(fncLogOutputTbl.DataBodyRange.rows.Count - 1, fncLogOutputTbl.DataBodyRange.columns.Count).rows.Delete
            For columns = 1 To fncLogOutputTbl.ListColumns("Tekst").Index
                fncLogOutputTbl.ListColumns(columns).DataBodyRange.Clear
            Next columns
        End If
    End If
    Set LogOutput = fncLogOutputTbl
    
End Function
Public Sub LogOutputLines(ByVal subTbl As ListObject, Optional ByVal subCtrlTxt As String, Optional ByVal subLogKey As String, Optional ByVal subLogTxt As String)
    
    Dim TblRows As Long
    Dim TblDate As Date
    Dim TblColumn As Long
    
    TblRows = subTbl.DataBodyRange.rows.Count
    TblDate = Now
    If Not subTbl.ListColumns("#").DataBodyRange(TblRows) = vbNullString Then
        subTbl.ListRows.Add
        TblRows = subTbl.DataBodyRange.rows.Count
    End If

    On Error Resume Next
    subTbl.AutoFilter.ShowAllData
    On Error GoTo 0
    
    For TblColumn = 1 To subTbl.ListColumns.Count
        If "#" = subTbl.ListColumns(TblColumn).Name Then subTbl.ListColumns("#").DataBodyRange(TblRows) = TblRows
        If "Dato" = subTbl.ListColumns(TblColumn).Name Then subTbl.ListColumns("Dato").DataBodyRange(TblRows) = CStr(TblDate)
        If "Handling" = subTbl.ListColumns(TblColumn).Name Then subTbl.ListColumns("Handling").DataBodyRange(TblRows) = subCtrlTxt
        If "N�gle" = subTbl.ListColumns(TblColumn).Name Then subTbl.ListColumns("N�gle").DataBodyRange(TblRows) = subLogKey
        If "Tekst" = subTbl.ListColumns(TblColumn).Name Then subTbl.ListColumns("Tekst").DataBodyRange(TblRows) = subLogTxt
    Next TblColumn
    
End Sub
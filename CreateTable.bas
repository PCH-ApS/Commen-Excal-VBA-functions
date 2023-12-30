Public Sub CreateTable(ByVal wksSub As Worksheet)
    
    Dim rngSub As range
    Dim tablename As String

    tablename = StrConv(wksSub.Name, vbUpperCase)
    If InStr(tablename, " ") > 0 Then tablename = Replace(tablename, " ", "_")

    Set rngSub = wksSub.UsedRange
    If wksSub.ListObjects.Count < 1 Then
        wksSub.ListObjects.Add(xlSrcRange, rngSub, , xlYes).Name = tablename
        wksSub.ListObjects(tablename).TableStyle = vbNullString
    End If
    
End Sub
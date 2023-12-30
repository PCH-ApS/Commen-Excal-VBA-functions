Public Sub CloseOpenWorkbook(filename As String)

    Dim closeWb As Workbook
    
    On Error Resume Next
    If Not filename = vbNullString Then
        Set closeWb = Workbooks.Open(filename, UpdateLinks:=0, Local:=True)
        If Not closeWb Is Nothing Then closeWb.Close (False)
    End If
    On Error GoTo 0

End Sub
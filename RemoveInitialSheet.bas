Public Sub RemoveInitialSheet(Optional ByVal subWb As Workbook)
    
    Application.DisplayAlerts = False
    Dim subWks As Worksheet
    If subWb Is Nothing Then
        For Each subWks In ThisWorkbook.Worksheets
            If subWks.CodeName = "Sheet1" Then
                subWks.Delete
            End If
        Next
    Else
        For Each subWks In subWb.Worksheets
            If subWks.CodeName = "Sheet1" Then
                subWks.Delete
            End If
        Next
    End If

    Application.DisplayAlerts = True
    
End Sub
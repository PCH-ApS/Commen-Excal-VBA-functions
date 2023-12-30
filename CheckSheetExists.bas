Public Function CheckSheetExists(ByVal NameToCheck As String, Optional ByVal WbToCheck As Workbook) As Boolean

    If WbToCheck Is Nothing Then
        Set WbToCheck = ThisWorkbook
    End If

    Dim WksToCheck As Worksheet
    On Error Resume Next
    Set WksToCheck = WbToCheck.Sheets(NameToCheck)
    On Error GoTo 0
    CheckSheetExists = Not WksToCheck Is Nothing
    
End Function
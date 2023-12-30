Public Sub RemoveTable(ByVal WbWithListObj As Workbook)
    
    Dim SheetWithListObj As Worksheet
    Dim listobj As ListObject
    Dim ListToRemove As Boolean
    
    For Each SheetWithListObj In WbWithListObj.Worksheets
        With SheetWithListObj
            ListToRemove = False
            For Each listobj In .ListObjects
                If StrConv(listobj.Name, vbUpperCase) = StrConv(.Name, vbUpperCase) Then ListToRemove = True
                If ListToRemove = True Then SheetWithListObj.ListObjects(StrConv(.Name, vbUpperCase)).Unlist
            Next listobj
        End With
    Next SheetWithListObj
    
End Sub
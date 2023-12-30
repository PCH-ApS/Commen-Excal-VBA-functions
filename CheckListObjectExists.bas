Public Function CheckListObjectExists(ByVal ListObjTCheck As String, ByVal WksToCheck As Worksheet) As Boolean
    
    Dim NewListObjTCheck As String
    If InStr(ListObjTCheck, " ") > 0 Then NewListObjTCheck = Replace(ListObjTCheck, " ", "_")
    If NewListObjTCheck = vbNullString Then NewListObjTCheck = ListObjTCheck
    Dim fncLo As ListObject
    On Error Resume Next
    Set fncLo = WksToCheck.ListObjects(StrConv(NewListObjTCheck, vbUpperCase))
    On Error GoTo 0
    CheckListObjectExists = Not fncLo Is Nothing
    
End Function
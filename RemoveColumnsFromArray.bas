Public Function RemoveColumnsFromArray(ByRef arr As Variant, colToKeep As Variant) As Variant
      
    Dim tempArr() As Variant
    ReDim tempArr(LBound(arr, 1) To UBound(arr, 1), LBound(colToKeep) To UBound(colToKeep)) As Variant
    
    Dim arrRows As Long
    Dim keepColumns As Long
    Dim arrColumns As Long
    
    For arrColumns = LBound(arr, 2) To UBound(arr, 2)
        For keepColumns = LBound(colToKeep) To UBound(colToKeep)
            If StrConv(CStr(arr(1, arrColumns)), vbUpperCase) = StrConv(CStr(colToKeep(keepColumns)), vbUpperCase) Then
                For arrRows = LBound(arr, 1) To UBound(arr, 1)
                    tempArr(arrRows, keepColumns) = arr(arrRows, arrColumns)
                Next arrRows
            End If
        Next keepColumns
    Next arrColumns
  
    RemoveColumnsFromArray = tempArr
    
End Function
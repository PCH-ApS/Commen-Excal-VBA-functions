Public Function RemoveEmptyRowsFromArray(ByRef arr As Variant) As Variant
    
    Dim i As Long
    Dim j As Long
    Dim tempIndex As Long
    
    'Redim preserve only works on the LAST dimention, so need to know how many data lines there are
    For i = LBound(arr, 1) To UBound(arr, 1)
    
        Dim allColumnsEmpty As Boolean
        allColumnsEmpty = True
        For j = LBound(arr, 2) To UBound(arr, 2)
            If Len(arr(i, j)) > 0 Then
                allColumnsEmpty = False
                Exit For
            End If
        Next j
        
        If Not allColumnsEmpty Then tempIndex = tempIndex + 1
        
    Next i
    
    'If any blank lines was found tempIndex will be smaller then the original array
    If tempIndex < UBound(arr, 1) Then
    
       Dim tempArr() As Variant
       ReDim tempArr(LBound(arr, 1) To tempIndex, LBound(arr, 2) To UBound(arr, 2))
       
       tempIndex = 0
       
       For i = LBound(arr, 1) To UBound(arr, 1)
           allColumnsEmpty = True
           For j = LBound(arr, 2) To UBound(arr, 2)
               If Len(arr(i, j)) > 0 Then
                   allColumnsEmpty = False
                   Exit For
               End If
           Next j
           
           If Not allColumnsEmpty Then
               tempIndex = tempIndex + 1
               
               For j = LBound(arr, 2) To UBound(arr, 2)
                   tempArr(tempIndex, j) = arr(i, j)
               Next j
           End If
            
       Next i
    End If
    
    If tempIndex = 0 Then
        'all rows were empty, so set array to empty
        Erase arr
    End If
    
    If tempIndex < UBound(arr, 1) Then
        RemoveEmptyRowsFromArray = tempArr
        Exit Function
    End If
  
    RemoveEmptyRowsFromArray = arr
    
End Function
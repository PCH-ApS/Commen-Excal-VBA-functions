Attribute VB_Name = "RemoveMatchingRowsFromArrayV2"
Option Explicit
'@Folder "Code Common"
Public Function RemoveMatchingRowsFromArrayV2(ByRef arr As Variant, ByVal lineIdent As String) As Variant

    Dim i As Long
    Dim j As Long
    Dim tempIndex As Long
    
    'Redim preserve only works on the LAST dimention, so need to know how many data lines there are
    For i = LBound(arr, 1) To UBound(arr, 1)
    
        Dim lineIdentFound As Boolean
        lineIdentFound = False
        For j = LBound(arr, 2) To UBound(arr, 2)
            If arr(i, j) = lineIdent Then
                lineIdentFound = True
                Exit For
            End If
        Next j
        
        If Not lineIdentFound Then tempIndex = tempIndex + 1
        
    Next i
    
    'If any lines with lineIdent was found tempIndex will be smaller then the original array
    If tempIndex < UBound(arr, 1) Then
    
       Dim tempArr() As Variant
       ReDim tempArr(LBound(arr, 1) To tempIndex, LBound(arr, 2) To UBound(arr, 2))
       
       tempIndex = 0
       
       For i = LBound(arr, 1) To UBound(arr, 1)
           lineIdentFound = False
           For j = LBound(arr, 2) To UBound(arr, 2)
               If arr(i, j) = lineIdent Then
                lineIdentFound = True
                Exit For
               End If
           Next j
           
           If Not lineIdentFound Then
               tempIndex = tempIndex + 1
               
               For j = LBound(arr, 2) To UBound(arr, 2)
                   tempArr(tempIndex, j) = arr(i, j)
               Next j
           End If
            
       Next i
    End If
    
    If tempIndex = 0 Or tempIndex = UBound(arr, 1) Then
        'No lineIdent was found, so exit
        RemoveMatchingRowsFromArrayV2 = arr
        Exit Function
    End If
    
    If tempIndex < UBound(arr, 1) Then
        RemoveMatchingRowsFromArrayV2 = tempArr
        Exit Function
    End If

End Function


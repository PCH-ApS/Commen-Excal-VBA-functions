Public Function AddMissingRowArrayToTargetArray(ByRef sourceArray As Variant, ByRef targetArray As Variant, headerArray As Variant, sourceRowIDarr As Variant) As Variant
    
    Dim newRowCount As Long
    newRowCount = 1
    Dim idArray() As Long
    Dim targetRowCount As Long
    For targetRowCount = LBound(targetArray, 1) To UBound(targetArray, 1)
        If targetArray(targetRowCount, 1) = vbNullString Then
            ReDim Preserve idArray(1 To newRowCount)
            idArray(newRowCount) = targetRowCount
            newRowCount = newRowCount + 1
        End If
    Next targetRowCount

    Dim targetRowID As Long
    If UBound(idArray) = UBound(sourceRowIDarr) Then
        Dim sourceRowID As Long
        Dim headerRowCount As Long
        For targetRowCount = LBound(idArray) To UBound(idArray)
            targetRowID = idArray(targetRowCount)
            sourceRowID = sourceRowIDarr(targetRowCount)
            For headerRowCount = LBound(headerArray, 1) To UBound(headerArray, 1)
                If Not InStr(headerArray(headerRowCount, 1), "*") > 0 Then
                    targetArray(targetRowID, headerArray(headerRowCount, 2)) = sourceArray(sourceRowID, headerArray(headerRowCount, 3))
                End If
            Next headerRowCount
        Next targetRowCount
    End If

    AddMissingRowArrayToTargetArray = targetArray

End Function
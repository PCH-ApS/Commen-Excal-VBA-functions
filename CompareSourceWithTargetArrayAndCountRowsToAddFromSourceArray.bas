Public Function CompareSourceWithTargetArrayAndCountRowsToAddFromSourceArray(ByRef sourceArray As Variant, ByRef targetArray As Variant, ByRef headerArray As Variant) As Variant

' **** Compare arrays and add rows from importArray
    Dim importRowCount As Long
    Dim sourceCompareString As String
    Dim importColumnCount As Long
    Dim headerRowCount As Long
    Dim newRowCount As Long
    newRowCount = 1
    Dim idArray() As Long
    ReDim Preserve idArray(1 To newRowCount)

' **** Count new rows to add to
    For importRowCount = LBound(sourceArray, 1) + 1 To UBound(sourceArray, 1)
        sourceCompareString = vbNullString
        For importColumnCount = LBound(sourceArray, 2) To UBound(sourceArray, 2)
            For headerRowCount = LBound(headerArray, 1) To UBound(headerArray, 1)
                If Not InStr(headerArray(headerRowCount, 1), "*") > 0 Then
                    If headerArray(headerRowCount, 1) = sourceArray(1, importColumnCount) Then
                        sourceCompareString = sourceCompareString & ";" & CStr(sourceArray(importRowCount, importColumnCount))
                        Exit For
                    End If
                End If
            Next headerRowCount
        Next importColumnCount
        If CommonFunctions.SourceSearchStringFoundInTargetArray(sourceCompareString, targetArray, headerArray) = False Then
            ReDim Preserve idArray(1 To newRowCount)
            idArray(newRowCount) = importRowCount
            newRowCount = newRowCount + 1
        End If
    Next importRowCount
    
    CompareSourceWithTargetArrayAndCountRowsToAddFromSourceArray = idArray

End Function
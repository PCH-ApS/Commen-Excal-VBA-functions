Public Function SourceSearchStringFoundInTargetArray(ByRef sourceSearchString As String, ByRef targetArray As Variant, headerArray As Variant) As Boolean
  
    Dim targetRowCount As Long
    Dim targetColumnCount As Long
    Dim headerRowCount As Long
    Dim targetCompareString As String

    For targetRowCount = LBound(targetArray, 1) + 1 To UBound(targetArray, 1)
        targetCompareString = vbNullString
        For targetColumnCount = LBound(targetArray, 2) To UBound(targetArray, 2)
            For headerRowCount = LBound(headerArray, 1) To UBound(headerArray, 1)
                If Not InStr(headerArray(headerRowCount, 1), "*") > 0 Then
                    If headerArray(headerRowCount, 1) = targetArray(1, targetColumnCount) Then
                        targetCompareString = targetCompareString & ";" & CStr(targetArray(targetRowCount, targetColumnCount))
                        Exit For
                    End If
                End If
            Next headerRowCount
        Next targetColumnCount
        If targetCompareString = sourceSearchString Then
            SourceSearchStringFoundInTargetArray = True
            Exit Function
        End If
    Next targetRowCount

    SourceSearchStringFoundInTargetArray = False

End Function
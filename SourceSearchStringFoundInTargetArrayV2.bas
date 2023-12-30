Public Function SourceSearchStringFoundInTargetArrayV2(ByRef sourceSearchString As String, ByRef targetArray As Variant, ByRef headerArray As Variant, ByVal searchKey As Variant) As Boolean
  
    Dim targetRowCount As Long
    Dim targetCompareString As String
    Dim targetColumnCount As Long
    Dim headerRowCount As Long
    
    For targetRowCount = LBound(targetArray, 1) + 1 To UBound(targetArray, 1)
        If targetArray(targetRowCount, 1) = searchKey Then
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
                SourceSearchStringFoundInTargetArrayV2 = True
                Exit Function
            End If
        End If
    Next targetRowCount
    
    SourceSearchStringFoundInTargetArrayV2 = False

End Function
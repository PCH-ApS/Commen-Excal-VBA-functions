Public Function GetColumnFromArray(ByRef arr As Variant, ByRef title As String) As Long

    Dim tableColumnLink As Long
    For tableColumnLink = LBound(arr, 2) To UBound(arr, 2)
        If arr(1, tableColumnLink) = title Then
            Exit For
        End If
    Next tableColumnLink

    GetColumnFromArray = tableColumnLink

End Function